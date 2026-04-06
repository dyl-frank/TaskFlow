"""
TaskFlow – Elegant Project Scheduler GUI
──────────────────────────────────────────
Built with CustomTkinter for a modern, native-feeling interface.
"""

from __future__ import annotations
import os, threading, datetime as dt, math
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
import yaml

import customtkinter as ctk
import networkx as nx

from scheduler import (
    load_yaml, build_graph, topological_order, schedule,
    export_gantt_xlsx, export_gantt_csv, export_ics,
    generate_calendar_links, ScheduleResult,
)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT  = "#4E79A7"
SUCCESS = "#59A14F"
WARNING = "#F28E2B"
ERROR   = "#E15759"
BG_DARK = "#1a1a2e"
BG_CARD = "#16213e"
BG_LIGHT = "#0f3460"

NODE_COLORS = [
    "#2563EB", "#059669", "#D97706", "#DC2626",
    "#7C3AED", "#0891B2", "#DB2777", "#65A30D",
    "#C2410C", "#0369A1",
]


# ── Helpers ────────────────────────────────────────────────────────────────

def _kahn_sort(ids: set[int], preds: dict[int, set[int]]) -> tuple[list[int], set[int]]:
    """Kahn's iterative topological sort.  Returns (topo_order, cycle_nodes)."""
    succs: dict[int, set[int]] = {i: set() for i in ids}
    for tid, ps in preds.items():
        for p in ps:
            if p in succs:
                succs[p].add(tid)
    in_deg = {tid: len(ps & ids) for tid, ps in preds.items()}
    queue = sorted(tid for tid, d in in_deg.items() if d == 0)
    order: list[int] = []
    while queue:
        tid = queue.pop(0)
        order.append(tid)
        for s in sorted(succs[tid]):
            in_deg[s] -= 1
            if in_deg[s] == 0:
                queue.append(s)
    cycles = ids - set(order)
    return order, cycles


def _compute_layers(ids: set[int], preds: dict[int, set[int]]) -> dict[int, int]:
    """Longest-path layer assignment (sources = layer 0)."""
    order, cycles = _kahn_sort(ids, preds)
    layer: dict[int, int] = {}
    for tid in order:
        if preds[tid] & ids:
            layer[tid] = max((layer.get(p, 0) for p in preds[tid] if p in ids), default=0) + 1
        else:
            layer[tid] = 0
    # Place cycle nodes at the end
    max_l = max(layer.values(), default=-1)
    for tid in sorted(cycles):
        max_l += 1
        layer[tid] = max_l
    return layer


# ── Node graph helper ───────────────────────────────────────────────────────

def _rrect_pts(x1: float, y1: float, x2: float, y2: float, r: float) -> list:
    """Polygon point list for a rounded rectangle (for tk.Canvas create_polygon)."""
    return [
        x1+r, y1,   x2-r, y1,
        x2,   y1,   x2,   y1+r,
        x2,   y2-r, x2,   y2,
        x2-r, y2,   x1+r, y2,
        x1,   y2,   x1,   y2-r,
        x1,   y1+r, x1,   y1,
    ]


class NodeGraph(tk.Canvas):
    """
    Interactive node-graph canvas (egui-style).
    Each task is a draggable node.  Drag from amber output port to blue input
    port to create a dependency connection.
    """

    NW      = 210    # node width
    TH      = 26     # title-bar height
    PR      = 7      # port circle radius
    ROW_H   = 17     # text row height inside body
    V_PAD   = 9      # vertical padding inside body
    N_ROWS  = 4      # description / areas / days / deps

    C = {
        "bg":       "#111827",
        "grid":     "#1a2535",
        "node":     "#1e293b",
        "node_sel": "#263450",
        "border":   "#374151",
        "sel_bdr":  "#F28E2B",
        "title_txt":"#f1f5f9",
        "body_txt": "#cbd5e1",
        "sub_txt":  "#64748b",
        "port_in":  "#60a5fa",   # blue  – receives connections
        "port_out": "#fbbf24",   # amber – sends connections
        "edge":     "#60a5fa",
        "edge_cyc": "#ef4444",
        "shadow":   "#050d1a",
    }

    TITLE_PALETTE = [
        "#1d4ed8", "#047857", "#b45309", "#be123c",
        "#6d28d9", "#0369a1", "#9d174d", "#3f6212",
        "#92400e", "#155e75",
    ]

    def __init__(self, parent,
                 on_select  = None,   # cb(tid: int)
                 on_connect = None,   # cb(from_tid: int, to_tid: int)  → add dep
                 on_remove  = None,   # cb(tid: int)
                 **kwargs):
        super().__init__(parent, bg=self.C["bg"], highlightthickness=0, **kwargs)
        self._on_select  = on_select
        self._on_connect = on_connect
        self._on_remove  = on_remove

        self._tasks: list[dict] = []
        self._pos:   dict[int, list[float]] = {}   # {tid: [x, y]}  – mutable list for easy update
        self._area_color: dict[str, str] = {}

        self._selected: int | None = None
        self._drag_id:  int | None = None
        self._drag_ox = 0.0;  self._drag_oy = 0.0   # offset within node

        # Port-connection drag state
        self._conn_src: int | None = None    # output port being dragged from
        self._conn_mx  = 0.0;  self._conn_my = 0.0

        self.bind("<ButtonPress-1>",   self._press)
        self.bind("<B1-Motion>",       self._motion)
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Configure>",       lambda _e: self._on_resize())

    # ── Public API ───────────────────────────────────────────────────────────

    def sync(self, tasks: list[dict], selected_id: int | None = None):
        """Sync canvas with current task list.  New tasks are auto-positioned."""
        self._tasks = tasks
        if selected_id is not None:
            self._selected = selected_id

        # Assign colors to new areas
        ci = len(self._area_color)
        for t in tasks:
            area = t["areas"][0] if t["areas"] else "?"
            if area not in self._area_color:
                self._area_color[area] = self.TITLE_PALETTE[ci % len(self.TITLE_PALETTE)]
                ci += 1

        # Position new tasks that don't have a position yet
        current_ids = {t["id"] for t in tasks}
        new_tasks   = [t for t in tasks if t["id"] not in self._pos]
        if new_tasks:
            self._place_new_nodes(new_tasks)

        # Drop stale positions
        self._pos = {k: v for k, v in self._pos.items() if k in current_ids}
        self._render()

    def set_selected(self, tid: int | None):
        self._selected = tid
        self._render()

    def auto_layout(self):
        """Re-compute positions for all nodes using topological layering."""
        if not self._tasks:
            return
        ids   = {t["id"] for t in self._tasks}
        preds = {t["id"]: {d for d in t["dependencies"] if d in ids} for t in self._tasks}
        layer = _compute_layers(ids, preds)
        groups: dict[int, list[int]] = {}
        for tid, l in layer.items():
            groups.setdefault(l, []).append(tid)

        cw = max(self.winfo_width(), 600)
        ch = max(self.winfo_height(), 400)
        n_layers = max(layer.values(), default=0) + 1
        nh = self._node_h()
        col_w = max(self.NW + 80, (cw - 40) / n_layers)

        for l_idx in range(n_layers):
            group = sorted(groups.get(l_idx, []))
            total = len(group) * (nh + 30) - 30
            sy = max(20, (ch - total) / 2)
            x = 20 + col_w * l_idx
            for i, tid in enumerate(group):
                self._pos[tid] = [x, sy + i * (nh + 30)]
        self._render()

    # ── Geometry ─────────────────────────────────────────────────────────────

    def _node_h(self) -> int:
        return self.TH + self.V_PAD * 2 + self.N_ROWS * self.ROW_H

    def _out_port(self, tid: int) -> tuple[float, float]:
        p = self._pos.get(tid, [0, 0])
        return (p[0] + self.NW, p[1] + self._node_h() / 2)

    def _in_port(self, tid: int) -> tuple[float, float]:
        p = self._pos.get(tid, [0, 0])
        return (p[0], p[1] + self._node_h() / 2)

    def _task_by_id(self, tid: int) -> dict | None:
        return next((t for t in self._tasks if t["id"] == tid), None)

    def _hit_node(self, x: float, y: float) -> int | None:
        """Return task id of topmost node under (x, y), or None."""
        nh = self._node_h()
        for t in reversed(self._tasks):   # reversed = top-most first
            tid = t["id"]
            if tid not in self._pos:
                continue
            nx_, ny = self._pos[tid]
            if nx_ <= x <= nx_ + self.NW and ny <= y <= ny + nh:
                return tid
        return None

    def _hit_in_port(self, x: float, y: float) -> int | None:
        for t in self._tasks:
            px, py = self._in_port(t["id"])
            if math.hypot(x - px, y - py) <= self.PR + 5:
                return t["id"]
        return None

    def _hit_out_port(self, x: float, y: float) -> int | None:
        for t in self._tasks:
            px, py = self._out_port(t["id"])
            if math.hypot(x - px, y - py) <= self.PR + 5:
                return t["id"]
        return None

    def _hit_close(self, tid: int, x: float, y: float) -> bool:
        p = self._pos.get(tid, [0, 0])
        bx = p[0] + self.NW - 14
        by = p[1] + self.TH / 2
        return math.hypot(x - bx, y - by) <= 9

    # ── Rendering ────────────────────────────────────────────────────────────

    def _render(self):
        self.delete("all")
        self._draw_grid()

        if not self._tasks:
            self.create_text(
                16, 16, anchor="nw",
                text="Add tasks using the form ← then drag output ports (●) to input ports (●) to create dependencies.",
                fill=self.C["sub_txt"], font=("Consolas", 10),
            )
            return

        ids   = {t["id"] for t in self._tasks}
        preds = {t["id"]: {d for d in t["dependencies"] if d in ids} for t in self._tasks}
        _, cycles = _kahn_sort(ids, preds)

        # Edges beneath nodes
        for t in self._tasks:
            for dep_id in t["dependencies"]:
                if dep_id in self._pos and t["id"] in self._pos:
                    in_cycle = t["id"] in cycles or dep_id in cycles
                    ox, oy = self._out_port(dep_id)
                    ix, iy = self._in_port(t["id"])
                    self._bezier(ox, oy, ix, iy,
                                 color=self.C["edge_cyc"] if in_cycle else self.C["edge"])

        # In-progress connection drag
        if self._conn_src is not None and self._conn_src in self._pos:
            ox, oy = self._out_port(self._conn_src)
            self._bezier(ox, oy, self._conn_mx, self._conn_my,
                         color=self.C["port_out"], dash=(7, 4))

        # Nodes on top
        for t in self._tasks:
            self._draw_node(t, cycles)

    def _draw_grid(self):
        cw, ch = max(self.winfo_width(), 100), max(self.winfo_height(), 100)
        step = 40
        for gx in range(0, cw + step, step):
            self.create_line(gx, 0, gx, ch, fill=self.C["grid"], width=1)
        for gy in range(0, ch + step, step):
            self.create_line(0, gy, cw, gy, fill=self.C["grid"], width=1)

    def _draw_node(self, task: dict, cycles: set[int]):
        tid = task["id"]
        if tid not in self._pos:
            return
        x, y = self._pos[tid]
        w, h  = self.NW, self._node_h()
        r     = 9
        is_sel = (tid == self._selected)
        in_cyc = (tid in cycles)

        bdr_col = self.C["edge_cyc"] if in_cyc else (self.C["sel_bdr"] if is_sel else self.C["border"])
        bdr_w   = 2.5 if (is_sel or in_cyc) else 1.0

        # Drop shadow
        sh = 4
        self.create_polygon(
            _rrect_pts(x+sh, y+sh, x+w+sh, y+h+sh, r),
            smooth=True, fill=self.C["shadow"], outline="",
        )

        # Node body
        self.create_polygon(
            _rrect_pts(x, y, x+w, y+h, r),
            smooth=True, fill=self.C["node_sel"] if is_sel else self.C["node"],
            outline=bdr_col, width=bdr_w,
            tags=f"nb_{tid}",
        )

        # Title bar – colored by first area
        area       = task["areas"][0] if task["areas"] else "?"
        title_col  = self._area_color.get(area, "#334155")
        # Full rounded rect for title bar then cover lower-rounded corners
        self.create_polygon(
            _rrect_pts(x, y, x+w, y+self.TH, r),
            smooth=True, fill=title_col, outline="",
            tags=f"nb_{tid}",
        )
        self.create_rectangle(
            x+1, y+self.TH-r-1, x+w-1, y+self.TH,
            fill=title_col, outline="",
            tags=f"nb_{tid}",
        )
        # Divider line between title and body
        self.create_line(x, y+self.TH, x+w, y+self.TH,
                         fill=bdr_col, width=bdr_w)

        # Title text
        self.create_text(
            x+10, y+self.TH/2, anchor="w",
            text=f"Task #{tid}", fill=self.C["title_txt"],
            font=("Consolas", 10, "bold"), tags=f"nb_{tid}",
        )
        # Close ✕
        self.create_text(
            x+w-12, y+self.TH/2, anchor="center",
            text="✕", fill="#94a3b8",
            font=("Consolas", 9), tags=f"close_{tid}",
        )

        # Body rows
        by = y + self.TH + self.V_PAD
        desc = task["description"]
        if len(desc) > 26: desc = desc[:24] + "…"
        self.create_text(
            x+14, by, anchor="w", text=desc,
            fill=self.C["body_txt"], font=("Consolas", 9, "bold"),
            tags=f"nb_{tid}",
        )
        by += self.ROW_H

        areas_s = ", ".join(task["areas"][:3]) or "—"
        self.create_text(
            x+14, by, anchor="w", text=f"areas:  {areas_s}",
            fill=self.C["sub_txt"], font=("Consolas", 9),
            tags=f"nb_{tid}",
        )
        by += self.ROW_H

        self.create_text(
            x+14, by, anchor="w", text=f"days:   {task['estimated_days']}",
            fill=self.C["sub_txt"], font=("Consolas", 9),
            tags=f"nb_{tid}",
        )
        by += self.ROW_H

        deps_s = ", ".join(str(d) for d in task["dependencies"]) or "none"
        if len(deps_s) > 22: deps_s = deps_s[:20] + "…"
        self.create_text(
            x+14, by, anchor="w", text=f"deps:   {deps_s}",
            fill=self.C["sub_txt"], font=("Consolas", 9),
            tags=f"nb_{tid}",
        )

        # ── Ports ──────────────────────────────────────────────────────
        nh2 = h / 2

        # Input port (left, blue)
        ix, iy = x, y + nh2
        self.create_oval(
            ix-self.PR, iy-self.PR, ix+self.PR, iy+self.PR,
            fill=self.C["port_in"], outline="#e2e8f0", width=1.5,
            tags=(f"pi_{tid}",),
        )
        self.create_text(
            ix+self.PR+5, iy, anchor="w",
            text="in", fill=self.C["port_in"],
            font=("Consolas", 8), tags=f"nb_{tid}",
        )

        # Output port (right, amber)
        ox_, oy_ = x+w, y + nh2
        self.create_oval(
            ox_-self.PR, oy_-self.PR, ox_+self.PR, oy_+self.PR,
            fill=self.C["port_out"], outline="#e2e8f0", width=1.5,
            tags=(f"po_{tid}",),
        )
        self.create_text(
            ox_-self.PR-5, oy_, anchor="e",
            text="out", fill=self.C["port_out"],
            font=("Consolas", 8), tags=f"nb_{tid}",
        )

        # Tag bindings
        self.tag_bind(f"nb_{tid}",    "<ButtonPress-1>",   lambda e, _t=tid: self._node_press(e, _t))
        self.tag_bind(f"close_{tid}", "<ButtonPress-1>",   lambda *_, _t=tid: self._close_press(_t))
        self.tag_bind(f"po_{tid}",    "<ButtonPress-1>",   lambda e, _t=tid: self._out_port_press(e, _t))
        self.tag_bind(f"pi_{tid}",    "<ButtonRelease-1>", lambda *_, _t=tid: self._in_port_release(_t))

    def _bezier(self, x1, y1, x2, y2, *, color="#60a5fa", dash=None):
        """Smooth cubic bezier polyline."""
        dx  = abs(x2 - x1)
        cp  = max(dx * 0.45, 70)
        cx1, cy1 = x1 + cp, y1
        cx2, cy2 = x2 - cp, y2
        pts = []
        N = 28
        for i in range(N + 1):
            t  = i / N
            mt = 1 - t
            pts.extend([
                mt**3*x1 + 3*mt**2*t*cx1 + 3*mt*t**2*cx2 + t**3*x2,
                mt**3*y1 + 3*mt**2*t*cy1 + 3*mt*t**2*cy2 + t**3*y2,
            ])
        kw: dict = {"fill": color, "width": 2.5, "smooth": False}
        if dash:
            kw["dash"] = dash
        self.create_line(*pts, **kw)
        # Arrowhead at target
        if len(pts) >= 4:
            ax, ay = pts[-2], pts[-1]
            bx, by = pts[-4], pts[-3]
            ux, uy = ax-bx, ay-by
            ln = math.hypot(ux, uy) or 1
            ux /= ln;  uy /= ln
            size = 9
            lx = ax - ux*size + (-uy)*size*0.45
            ly = ay - uy*size + ux*size*0.45
            rx = ax - ux*size - (-uy)*size*0.45
            ry = ay - uy*size - ux*size*0.45
            self.create_polygon([ax,ay, lx,ly, rx,ry], fill=color, outline="")

    # ── Mouse events ─────────────────────────────────────────────────────────

    def _node_press(self, event, tid: int):
        """Clicked on a node body – select + start drag."""
        if self._conn_src is not None:
            return    # mid-connection drag
        self._drag_id = tid
        p = self._pos[tid]
        self._drag_ox = event.x - p[0]
        self._drag_oy = event.y - p[1]
        if self._selected != tid:
            self._selected = tid
            if self._on_select:
                self._on_select(tid)
            self._render()

    def _out_port_press(self, event, tid: int):
        """Clicked on an output port – start connection drag."""
        self._conn_src = tid
        self._conn_mx  = event.x
        self._conn_my  = event.y
        self._drag_id  = None
        return "break"

    def _in_port_release(self, tid: int):
        """Released on an input port – complete connection."""
        if self._conn_src is not None and self._conn_src != tid:
            if self._on_connect:
                self._on_connect(self._conn_src, tid)
        self._conn_src = None
        self._render()

    def _close_press(self, tid: int):
        """Clicked the ✕ on a node."""
        if self._on_remove:
            self._on_remove(tid)

    def _press(self, event):
        """Canvas background press."""
        if self._conn_src is not None:
            # Check if we hit an input port
            hit = self._hit_in_port(event.x, event.y)
            if hit is not None and hit != self._conn_src:
                if self._on_connect:
                    self._on_connect(self._conn_src, hit)
            self._conn_src = None
            self._render()

    def _motion(self, event):
        if self._conn_src is not None:
            self._conn_mx = event.x
            self._conn_my = event.y
            self._render()
            return
        if self._drag_id is not None and self._drag_id in self._pos:
            self._pos[self._drag_id][0] = event.x - self._drag_ox
            self._pos[self._drag_id][1] = event.y - self._drag_oy
            self._render()

    def _release(self, event):
        if self._conn_src is not None:
            hit = self._hit_in_port(event.x, event.y)
            if hit is not None and hit != self._conn_src:
                if self._on_connect:
                    self._on_connect(self._conn_src, hit)
            self._conn_src = None
            self._render()
        self._drag_id = None

    # ── Internal ─────────────────────────────────────────────────────────────

    def _on_resize(self):
        # On first resize, auto-layout if tasks exist but no positions set
        if self._tasks and not self._pos:
            self.auto_layout()
        else:
            self._render()

    def _place_new_nodes(self, new_tasks: list[dict]):
        """Auto-position newly added tasks."""
        if not self._pos:
            # First batch → full layout via auto_layout logic
            ids   = {t["id"] for t in self._tasks}
            preds = {t["id"]: {d for d in t["dependencies"] if d in ids} for t in self._tasks}
            layer = _compute_layers(ids, preds)
            groups: dict[int, list[int]] = {}
            for tid, l in layer.items():
                groups.setdefault(l, []).append(tid)
            cw = max(self.winfo_width(), 600)
            ch = max(self.winfo_height(), 400)
            n_layers = max(layer.values(), default=0) + 1
            nh = self._node_h()
            col_w = max(self.NW + 80, (cw - 40) / n_layers)
            for l_idx in range(n_layers):
                group = sorted(groups.get(l_idx, []))
                total = len(group) * (nh + 30) - 30
                sy = max(20, (ch - total) / 2)
                x = 20 + col_w * l_idx
                for i, tid in enumerate(group):
                    if tid not in self._pos:
                        self._pos[tid] = [x, sy + i * (nh + 30)]
        else:
            # Append new tasks to the right
            max_x = max(p[0] for p in self._pos.values())
            cw    = max(self.winfo_width(), 600)
            x = min(max_x + self.NW + 70, cw - self.NW - 20)
            nh = self._node_h()
            for i, t in enumerate(new_tasks):
                self._pos[t["id"]] = [x, 30 + i * (nh + 30)]


class TaskFlowApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.title("TaskFlow – Project Scheduler")
        self.geometry("1380x860")
        self.minsize(1000, 680)

        self._result: ScheduleResult | None = None
        self._yaml_path: str | None = None

        # Editor state
        self._editor_tasks: list[dict] = []   # {id, description, areas, dependencies, estimated_days}
        self._project_start  = dt.date.today()
        self._project_end    = dt.date.today() + dt.timedelta(days=30)
        self._max_tasks      = 2
        self._members_raw: list[dict] = []    # raw dicts for YAML round-trip
        self._editing_id: int | None = None   # task ID currently loaded in form

        self._build_ui()

    # ── Layout ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Top banner
        banner = ctk.CTkFrame(self, height=56, corner_radius=0, fg_color=ACCENT)
        banner.pack(fill="x")
        banner.pack_propagate(False)
        ctk.CTkLabel(
            banner, text="⬡  TaskFlow Scheduler",
            font=ctk.CTkFont(size=22, weight="bold"), text_color="white",
        ).pack(side="left", padx=20)
        self._theme_btn = ctk.CTkButton(
            banner, text="☀ Light", width=80, height=30,
            fg_color="transparent", border_width=1, border_color="white",
            text_color="white", hover_color="#3a6a9a",
            command=self._toggle_theme,
        )
        self._theme_btn.pack(side="right", padx=20)

        # Main container
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=16, pady=(12, 16))
        container.grid_columnconfigure(1, weight=1)
        container.grid_rowconfigure(0, weight=1)

        # ── Left panel ──────────────────────────────────────────────────
        left = ctk.CTkFrame(container, width=310, corner_radius=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.grid_propagate(False)

        ctk.CTkLabel(
            left, text="Input Configuration",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(pady=(20, 6), padx=20, anchor="w")
        ctk.CTkFrame(left, height=2, fg_color=ACCENT).pack(fill="x", padx=20, pady=(0, 16))

        # File selector
        file_frame = ctk.CTkFrame(left, fg_color="transparent")
        file_frame.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(file_frame, text="YAML File:", font=ctk.CTkFont(size=12)).pack(anchor="w")
        btn_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=4)
        self._file_label = ctk.CTkLabel(
            btn_frame, text="No file selected",
            font=ctk.CTkFont(size=11), text_color="gray",
            wraplength=200, anchor="w",
        )
        self._file_label.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(
            btn_frame, text="Browse…", width=90, height=32,
            command=self._browse_file,
        ).pack(side="right")

        # Export format
        ctk.CTkLabel(left, text="Export Format:", font=ctk.CTkFont(size=12)).pack(padx=20, anchor="w", pady=(16, 4))
        self._format_var = ctk.StringVar(value="xlsx")
        fmt_frame = ctk.CTkFrame(left, fg_color="transparent")
        fmt_frame.pack(padx=20, anchor="w")
        ctk.CTkRadioButton(fmt_frame, text="Excel (.xlsx)", variable=self._format_var, value="xlsx").pack(side="left", padx=(0, 12))
        ctk.CTkRadioButton(fmt_frame, text="CSV (.csv)",   variable=self._format_var, value="csv").pack(side="left", padx=(0, 12))
        fmt_frame2 = ctk.CTkFrame(left, fg_color="transparent")
        fmt_frame2.pack(padx=20, anchor="w", pady=(4, 0))
        ctk.CTkRadioButton(fmt_frame2, text="Calendar (.ics)", variable=self._format_var, value="ics").pack(side="left")

        # Action buttons
        btn_container = ctk.CTkFrame(left, fg_color="transparent")
        btn_container.pack(fill="x", padx=20, pady=(24, 0))

        self._run_btn = ctk.CTkButton(
            btn_container, text="▶  Run Scheduler", height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=SUCCESS, hover_color="#4a9142",
            command=self._run_scheduler,
        )
        self._run_btn.pack(fill="x", pady=(0, 8))

        self._export_btn = ctk.CTkButton(
            btn_container, text="💾  Export Gantt Chart", height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=ACCENT, hover_color="#3d6a96",
            command=self._export, state="disabled",
        )
        self._export_btn.pack(fill="x", pady=(0, 8))

        self._cal_btn = ctk.CTkButton(
            btn_container, text="📅  Share to Calendar", height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#7C3AED", hover_color="#6D28D9",
            command=self._open_calendar_dialog, state="disabled",
        )
        self._cal_btn.pack(fill="x")

        self._progress = ctk.CTkProgressBar(left, height=6)
        self._progress.pack(fill="x", padx=20, pady=(20, 0))
        self._progress.set(0)

        self._status_label = ctk.CTkLabel(left, text="Ready", font=ctk.CTkFont(size=11), text_color="gray")
        self._status_label.pack(padx=20, pady=(6, 0), anchor="w")

        self._summary_frame = ctk.CTkFrame(left, corner_radius=10, fg_color=BG_LIGHT)
        self._summary_frame.pack(fill="x", padx=20, pady=(20, 20))
        self._summary_label = ctk.CTkLabel(
            self._summary_frame, text="Load a YAML file to begin.",
            font=ctk.CTkFont(size=11), justify="left", wraplength=250,
        )
        self._summary_label.pack(padx=14, pady=14)

        # ── Right panel (tabs) ───────────────────────────────────────────
        right = ctk.CTkFrame(container, corner_radius=12)
        right.grid(row=0, column=1, sticky="nsew")

        self._tabview = ctk.CTkTabview(right, corner_radius=10)
        self._tabview.pack(fill="both", expand=True, padx=10, pady=10)

        self._tab_schedule = self._tabview.add("📋 Schedule")
        self._tab_graph    = self._tabview.add("🔗 Dependencies")
        self._tab_editor   = self._tabview.add("🧩 Task Editor")
        self._tab_calendar = self._tabview.add("📅 Calendar")
        self._tab_log      = self._tabview.add("📝 Log")

        # Schedule tab
        self._schedule_text = ctk.CTkTextbox(
            self._tab_schedule, font=ctk.CTkFont(family="Consolas", size=12), wrap="none",
        )
        self._schedule_text.pack(fill="both", expand=True, padx=4, pady=4)

        # Dependencies tab
        graph_split = ctk.CTkFrame(self._tab_graph)
        graph_split.pack(fill="both", expand=True, padx=4, pady=4)
        graph_split.grid_rowconfigure(0, weight=1)
        graph_split.grid_columnconfigure(0, weight=1)
        graph_split.grid_columnconfigure(1, weight=1)

        self._graph_canvas = tk.Canvas(graph_split, bg="#111827", highlightthickness=0)
        self._graph_canvas.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=2)
        self._graph_canvas.create_text(
            10, 10, anchor="nw",
            text="Task Flow Visual: Run scheduler to render graph.",
            fill="white",
        )

        self._graph_text = ctk.CTkTextbox(
            graph_split, font=ctk.CTkFont(family="Consolas", size=12),
        )
        self._graph_text.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=2)

        # Task Editor tab
        self._build_task_editor_tab()

        # Calendar tab
        cal_header = ctk.CTkFrame(self._tab_calendar, fg_color="transparent")
        cal_header.pack(fill="x", padx=8, pady=(8, 0))
        ctk.CTkLabel(cal_header, text="Calendar Integration", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        self._cal_quick_export = ctk.CTkButton(
            cal_header, text="⬇ Quick Export .ics", width=150, height=30,
            fg_color="#7C3AED", hover_color="#6D28D9",
            command=self._quick_export_ics, state="disabled",
        )
        self._cal_quick_export.pack(side="right")
        self._cal_text = ctk.CTkTextbox(self._tab_calendar, font=ctk.CTkFont(family="Consolas", size=11))
        self._cal_text.pack(fill="both", expand=True, padx=4, pady=4)

        # Log tab
        self._log_text = ctk.CTkTextbox(self._tab_log, font=ctk.CTkFont(family="Consolas", size=11))
        self._log_text.pack(fill="both", expand=True, padx=4, pady=4)

    def _build_task_editor_tab(self):
        """Build the full Task Editor tab with form, task list, and flow diagram."""
        outer = ctk.CTkFrame(self._tab_editor, fg_color="transparent")
        outer.pack(fill="both", expand=True, padx=6, pady=6)
        outer.grid_columnconfigure(0, weight=3)
        outer.grid_columnconfigure(1, weight=5)
        outer.grid_rowconfigure(0, weight=1)

        # ── Left column: form + task list ─────────────────────────────────
        left_col = ctk.CTkFrame(outer, corner_radius=8)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left_col.grid_rowconfigure(3, weight=1)
        left_col.grid_columnconfigure(0, weight=1)

        # Form header
        hdr = ctk.CTkFrame(left_col, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
        ctk.CTkLabel(hdr, text="Add / Edit Task", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        self._autosave_label = ctk.CTkLabel(hdr, text="", font=ctk.CTkFont(size=10), text_color=SUCCESS)
        self._autosave_label.pack(side="right")

        # Form
        form = ctk.CTkFrame(left_col, corner_radius=6)
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 6))
        form.grid_columnconfigure(1, weight=1)

        def lbl(parent, text, row):
            ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=12)).grid(
                row=row, column=0, sticky="nw", padx=8, pady=5)

        lbl(form, "Task ID:", 0)
        self._task_id_var = tk.StringVar()
        ctk.CTkEntry(form, textvariable=self._task_id_var, width=70).grid(row=0, column=1, sticky="w", padx=8, pady=5)

        lbl(form, "Description:", 1)
        self._task_desc_var = tk.StringVar()
        ctk.CTkEntry(form, textvariable=self._task_desc_var).grid(row=1, column=1, sticky="ew", padx=8, pady=5)

        lbl(form, "Areas:", 2)
        self._task_areas_var = tk.StringVar()
        ctk.CTkEntry(form, textvariable=self._task_areas_var, placeholder_text="e.g. TH, materials").grid(
            row=2, column=1, sticky="ew", padx=8, pady=5)

        lbl(form, "Duration (days):", 3)
        self._task_duration_var = tk.StringVar(value="1")
        ctk.CTkEntry(form, textvariable=self._task_duration_var, width=70).grid(row=3, column=1, sticky="w", padx=8, pady=5)

        lbl(form, "Dependencies:", 4)
        dep_wrap = tk.Frame(form, bg="#1e1e2e")
        dep_wrap.grid(row=4, column=1, sticky="ew", padx=8, pady=5)
        self._dep_listbox = tk.Listbox(
            dep_wrap, selectmode=tk.MULTIPLE, height=5,
            bg="#1e1e2e", fg="white",
            selectbackground=ACCENT, selectforeground="white",
            font=("Consolas", 10), relief="flat", bd=1,
            highlightcolor=ACCENT, highlightthickness=1,
            activestyle="none",
        )
        dep_sb = tk.Scrollbar(dep_wrap, orient="vertical", command=self._dep_listbox.yview)
        self._dep_listbox.config(yscrollcommand=dep_sb.set)
        self._dep_listbox.pack(side="left", fill="both", expand=True)
        dep_sb.pack(side="right", fill="y")

        ctk.CTkLabel(form, text="Ctrl+click for multiple", font=ctk.CTkFont(size=9), text_color="gray").grid(
            row=5, column=1, sticky="w", padx=8, pady=(0, 4))

        # Action buttons
        btn_row = ctk.CTkFrame(left_col, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 8))
        ctk.CTkButton(btn_row, text="✚ Add / Update", width=120, fg_color=SUCCESS,
                      command=self._add_task).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_row, text="✖ Remove",       width=90,  fg_color=ERROR,
                      command=self._remove_task).pack(side="left", padx=4)
        ctk.CTkButton(btn_row, text="↺ Clear",        width=70,  fg_color="gray40",
                      command=self._clear_task_form).pack(side="left", padx=4)
        ctk.CTkButton(btn_row, text="💾 Save YAML",   width=100, fg_color=ACCENT,
                      command=self._save_current_yaml).pack(side="right")

        # Task list
        list_hdr = ctk.CTkFrame(left_col, fg_color="transparent")
        list_hdr.grid(row=3, column=0, sticky="new", padx=10, pady=(0, 2))
        ctk.CTkLabel(list_hdr, text="Tasks", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkLabel(list_hdr, text="(dep-sorted)", font=ctk.CTkFont(size=10), text_color="gray").pack(side="left", padx=6)

        task_list_wrap = tk.Frame(left_col, bg="#1e1e2e")
        task_list_wrap.grid(row=3, column=0, sticky="nsew", padx=10, pady=(20, 8))
        self._task_listbox = tk.Listbox(
            task_list_wrap, height=14,
            bg="#1e1e2e", fg="white",
            selectbackground=ACCENT, selectforeground="white",
            font=("Consolas", 10), relief="flat", bd=1,
            highlightcolor=ACCENT, highlightthickness=1,
            activestyle="none",
        )
        tl_sb = tk.Scrollbar(task_list_wrap, orient="vertical", command=self._task_listbox.yview)
        self._task_listbox.config(yscrollcommand=tl_sb.set)
        self._task_listbox.pack(side="left", fill="both", expand=True)
        tl_sb.pack(side="right", fill="y")
        self._task_listbox.bind("<<ListboxSelect>>", self._load_selected_task)

        # ── Right column: flow diagram ─────────────────────────────────────
        flow_card = ctk.CTkFrame(outer, corner_radius=8)
        flow_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        flow_card.grid_rowconfigure(1, weight=1)
        flow_card.grid_columnconfigure(0, weight=1)

        flow_hdr = ctk.CTkFrame(flow_card, fg_color="transparent")
        flow_hdr.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        ctk.CTkLabel(flow_hdr, text="Task Flow Diagram", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        self._cycle_label = ctk.CTkLabel(flow_hdr, text="", font=ctk.CTkFont(size=11))
        self._cycle_label.pack(side="left", padx=10)
        ctk.CTkButton(
            flow_hdr, text="⟳ Auto Layout", width=100, height=24,
            fg_color="gray30", hover_color="gray40",
            command=self._auto_layout_nodes,
        ).pack(side="right", padx=(4, 0))
        ctk.CTkLabel(flow_hdr, text="drag ● → ● to connect", font=ctk.CTkFont(size=9), text_color="gray").pack(side="right", padx=8)
        ctk.CTkLabel(flow_hdr, text="Order:", font=ctk.CTkFont(size=10), text_color="gray").pack(side="left", padx=(6, 2))
        self._topo_label = ctk.CTkLabel(
            flow_hdr, text="—",
            font=ctk.CTkFont(family="Consolas", size=10), text_color="#a0a0a0",
            wraplength=380, justify="left",
        )
        self._topo_label.pack(side="left")

        self._node_graph = NodeGraph(
            flow_card,
            on_select  = self._on_ng_select,
            on_connect = self._on_ng_connect,
            on_remove  = self._on_ng_remove,
        )
        self._node_graph.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))

    # ── Task editor logic ───────────────────────────────────────────────────

    def _clear_task_form(self):
        self._editing_id = None
        self._task_id_var.set("")
        self._task_desc_var.set("")
        self._task_areas_var.set("")
        self._task_duration_var.set("1")
        self._dep_listbox.selection_clear(0, tk.END)

    def _refresh_task_editor(self, *, auto_save: bool = False):
        """Rebuild the task list and dep listbox; optionally auto-save YAML."""
        ids = {t["id"] for t in self._editor_tasks}
        preds = {t["id"]: set(t["dependencies"]) for t in self._editor_tasks}
        order, cycles = _kahn_sort(ids, preds)
        # tasks not in order (cycles) go at the end
        remaining = [t["id"] for t in self._editor_tasks if t["id"] not in order]
        sorted_ids = order + remaining
        task_by_id = {t["id"]: t for t in self._editor_tasks}

        # Rebuild task listbox
        self._task_listbox.delete(0, tk.END)
        self._listbox_id_order: list[int] = []
        for tid in sorted_ids:
            if tid not in task_by_id:
                continue
            t = task_by_id[tid]
            deps_str = ",".join(str(d) for d in t["dependencies"]) if t["dependencies"] else "—"
            cycle_mark = " ⚠" if tid in cycles else ""
            label = f"#{tid:<3} {t['description'][:22]:<22}  deps:[{deps_str}]  {t['estimated_days']}d{cycle_mark}"
            self._task_listbox.insert(tk.END, label)
            self._listbox_id_order.append(tid)

        # Rebuild dependency listbox (all tasks except the one being edited)
        current_sel_ids = set()
        for i in self._dep_listbox.curselection():
            if i < len(self._dep_listbox_id_order):
                current_sel_ids.add(self._dep_listbox_id_order[i])

        self._dep_listbox.delete(0, tk.END)
        self._dep_listbox_id_order: list[int] = []
        for tid in sorted_ids:
            if tid not in task_by_id:
                continue
            if tid == self._editing_id:
                continue  # can't depend on yourself
            t = task_by_id[tid]
            label = f"#{tid}: {t['description'][:28]}"
            self._dep_listbox.insert(tk.END, label)
            self._dep_listbox_id_order.append(tid)

        # Re-apply selections based on currently edited task's deps
        if self._editing_id is not None and self._editing_id in task_by_id:
            existing_deps = set(task_by_id[self._editing_id]["dependencies"])
        else:
            existing_deps = current_sel_ids  # preserve user selection while typing
        for i, tid in enumerate(self._dep_listbox_id_order):
            if tid in existing_deps:
                self._dep_listbox.selection_set(i)

        if auto_save and self._yaml_path:
            self._do_save_yaml(self._yaml_path, silent=True)

        self._draw_task_flow()

    def _add_task(self):
        """Add a new task or update the one being edited."""
        raw_id = self._task_id_var.get().strip()
        desc   = self._task_desc_var.get().strip()
        areas  = [a.strip() for a in self._task_areas_var.get().split(",") if a.strip()]
        try:
            duration = int(self._task_duration_var.get().strip())
        except ValueError:
            messagebox.showerror("Invalid Input", "Duration must be an integer number of days.")
            return
        if not raw_id or not desc:
            messagebox.showerror("Invalid Input", "Task ID and Description are required.")
            return
        try:
            tid = int(raw_id)
        except ValueError:
            messagebox.showerror("Invalid Input", "Task ID must be an integer.")
            return

        # Read selected dependencies
        sel_indices = list(self._dep_listbox.curselection())
        dep_ids = [self._dep_listbox_id_order[i] for i in sel_indices
                   if i < len(self._dep_listbox_id_order)]

        # Check for duplicate ID (when adding new)
        existing = {t["id"]: t for t in self._editor_tasks}
        if tid in existing and self._editing_id != tid:
            messagebox.showerror("Duplicate ID", f"Task #{tid} already exists. Load it to edit.")
            return

        task_dict = {
            "id": tid,
            "description": desc,
            "areas": areas,
            "dependencies": dep_ids,
            "estimated_days": duration,
        }
        if tid in existing:
            idx = next(i for i, t in enumerate(self._editor_tasks) if t["id"] == tid)
            self._editor_tasks[idx] = task_dict
            self._log(f"Updated task #{tid}: {desc}")
        else:
            self._editor_tasks.append(task_dict)
            self._log(f"Added task #{tid}: {desc}")

        self._editing_id = tid
        self._refresh_task_editor(auto_save=bool(self._yaml_path))
        self._flash_autosave()

    def _remove_task(self):
        """Remove the task currently selected in the task listbox."""
        sel = self._task_listbox.curselection()
        if not sel:
            messagebox.showwarning("Nothing Selected", "Click a task in the list to select it first.")
            return
        idx = sel[0]
        if idx >= len(self._listbox_id_order):
            return
        tid = self._listbox_id_order[idx]
        task_by_id = {t["id"]: t for t in self._editor_tasks}
        desc = task_by_id.get(tid, {}).get("description", str(tid))
        if not messagebox.askyesno("Remove Task", f"Remove task #{tid}: {desc}?"):
            return

        # Remove task and any references to it in other tasks' deps
        self._editor_tasks = [t for t in self._editor_tasks if t["id"] != tid]
        for t in self._editor_tasks:
            t["dependencies"] = [d for d in t["dependencies"] if d != tid]

        self._log(f"Removed task #{tid}")
        self._clear_task_form()
        self._refresh_task_editor(auto_save=bool(self._yaml_path))
        self._flash_autosave()

    def _load_selected_task(self, _event=None):
        """Load the selected task from the listbox into the form."""
        sel = self._task_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self._listbox_id_order):
            return
        tid = self._listbox_id_order[idx]
        task_by_id = {t["id"]: t for t in self._editor_tasks}
        t = task_by_id.get(tid)
        if not t:
            return

        self._editing_id = tid
        self._task_id_var.set(str(t["id"]))
        self._task_desc_var.set(t["description"])
        self._task_areas_var.set(", ".join(t["areas"]))
        self._task_duration_var.set(str(t["estimated_days"]))
        # Refresh dep listbox with correct selection
        self._refresh_task_editor()
        self._draw_task_flow()

    def _select_task_by_id(self, tid: int):
        """Select a task in the listbox by ID (from canvas click)."""
        if not hasattr(self, "_listbox_id_order"):
            return
        if tid in self._listbox_id_order:
            idx = self._listbox_id_order.index(tid)
            self._task_listbox.selection_clear(0, tk.END)
            self._task_listbox.selection_set(idx)
            self._task_listbox.see(idx)
            self._load_selected_task()

    def _get_selected_task_id(self) -> int | None:
        return self._editing_id

    def _flash_autosave(self):
        self._autosave_label.configure(text="✓ Saved" if self._yaml_path else "✓ Updated")
        self.after(2500, lambda: self._autosave_label.configure(text=""))

    def _save_current_yaml(self):
        if not self._editor_tasks:
            messagebox.showwarning("No Tasks", "Add some tasks first.")
            return
        if self._yaml_path:
            path = self._yaml_path
        else:
            path = filedialog.asksaveasfilename(
                title="Save YAML",
                defaultextension=".yaml",
                filetypes=[("YAML files", "*.yaml *.yml"), ("All files", "*.*")],
                initialfile="tasks.yaml",
            )
            if not path:
                return
            self._yaml_path = path
            self._file_label.configure(text=os.path.basename(path), text_color=SUCCESS)

        self._do_save_yaml(path, silent=False)

    def _do_save_yaml(self, path: str, *, silent: bool = False):
        """Write editor tasks + project settings to a YAML file."""
        data = {
            "start_date": self._project_start.strftime("%m/%d/%Y"),
            "end_date":   self._project_end.strftime("%m/%d/%Y"),
            "maximum_tasks": self._max_tasks,
            "members": self._members_raw if self._members_raw else [
                {"name": "Unassigned", "areas": [], "dates_unavailable": []}
            ],
            "tasks": [
                {
                    "ID": t["id"],
                    "description": t["description"],
                    "areas": t["areas"],
                    "dependencies": t["dependencies"],
                    "estimated_time_days": t["estimated_days"],
                }
                for t in self._editor_tasks
            ],
        }
        try:
            with open(path, "w") as f:
                yaml.dump(data, f, default_flow_style=False, sort_keys=False, allow_unicode=True)
            self._log(f"YAML saved: {path}")
            if not silent:
                self._set_status(f"✓ Saved: {os.path.basename(path)}", SUCCESS)
                messagebox.showinfo("Saved", f"YAML saved to:\n{path}")
            self._flash_autosave()
        except Exception as e:
            self._log(f"YAML save error: {e}")
            if not silent:
                messagebox.showerror("Save Error", str(e))

    # ── Flow diagram ────────────────────────────────────────────────────────

    def _draw_task_flow(self):
        """Sync node graph + update header labels."""
        tasks = self._editor_tasks
        ids   = {t["id"] for t in tasks}
        preds = {t["id"]: {d for d in t["dependencies"] if d in ids} for t in tasks}
        order, cycles = _kahn_sort(ids, preds)

        topo_str = " → ".join(str(t) for t in order)
        if tasks:
            if cycles:
                cycle_str = ", ".join(str(c) for c in sorted(cycles))
                self._cycle_label.configure(text=f"⚠ Cycles: {cycle_str}", text_color=ERROR)
            else:
                self._cycle_label.configure(text="✓ No cycles", text_color=SUCCESS)
        else:
            self._cycle_label.configure(text="")
        self._topo_label.configure(text=topo_str or "—")

        self._node_graph.sync(tasks, selected_id=self._editing_id)

    # ── Node graph callbacks ─────────────────────────────────────────────────

    def _on_ng_select(self, tid: int):
        self._select_task_by_id(tid)

    def _on_ng_connect(self, from_tid: int, to_tid: int):
        """Drag from output of from_tid to input of to_tid → to_tid depends on from_tid."""
        task = next((t for t in self._editor_tasks if t["id"] == to_tid), None)
        if task is None:
            return
        if from_tid in task["dependencies"]:
            return   # already exists
        task["dependencies"].append(from_tid)
        self._log(f"Connected: task #{to_tid} now depends on #{from_tid}")
        self._refresh_task_editor(auto_save=bool(self._yaml_path))

    def _on_ng_remove(self, tid: int):
        """Node ✕ clicked → confirm + remove task."""
        task_by_id = {t["id"]: t for t in self._editor_tasks}
        desc = task_by_id.get(tid, {}).get("description", str(tid))
        if not messagebox.askyesno("Remove Task", f"Remove task #{tid}: {desc}?"):
            return
        self._editor_tasks = [t for t in self._editor_tasks if t["id"] != tid]
        for t in self._editor_tasks:
            t["dependencies"] = [d for d in t["dependencies"] if d != tid]
        if self._editing_id == tid:
            self._clear_task_form()
        self._log(f"Removed task #{tid}")
        self._refresh_task_editor(auto_save=bool(self._yaml_path))

    def _auto_layout_nodes(self):
        self._node_graph.auto_layout()

    def _canvas_arrow(self, canvas, x1, y1, x2, y2, *, in_cycle=False):
        """Draw a curved directed arrow on canvas (used by Dependencies tab)."""
        color = "#FF5555" if in_cycle else "#7BAFD4"
        dx = x2 - x1
        cx1 = x1 + dx * 0.35;  cy1 = y1
        cx2 = x1 + dx * 0.65;  cy2 = y2
        pts = []
        for i in range(21):
            t  = i / 20;  mt = 1 - t
            pts.extend([
                mt**3*x1 + 3*mt**2*t*cx1 + 3*mt*t**2*cx2 + t**3*x2,
                mt**3*y1 + 3*mt**2*t*cy1 + 3*mt*t**2*cy2 + t**3*y2,
            ])
        canvas.create_line(*pts, fill=color, width=2, smooth=False)
        dx2 = x2 - pts[-4];  dy2 = y2 - pts[-3]
        ln  = math.hypot(dx2, dy2) or 1
        ux, uy = dx2/ln, dy2/ln;  px, py = -uy, ux
        size = 8
        canvas.create_polygon(
            [x2, y2,
             x2 - ux*size + px*(size/2), y2 - uy*size + py*(size/2),
             x2 - ux*size - px*(size/2), y2 - uy*size - py*(size/2)],
            fill=color, outline=""
        )

    # ── Scheduler flow drawing (Dependencies tab) ────────────────────────────

    def _draw_graph_canvas(self, result: ScheduleResult):
        """Render the scheduler result graph on the Dependencies tab canvas."""
        canvas = self._graph_canvas
        canvas.delete("all")

        tasks_by_id = {t.id: t for t in result.tasks}
        ids   = set(tasks_by_id.keys())
        preds = {t.id: set(t.dependencies) for t in result.tasks}
        layer = _compute_layers(ids, preds)

        cw = canvas.winfo_width()
        ch = canvas.winfo_height()
        if cw < 20: cw = 500
        if ch < 20: ch = 400

        BOX_W, BOX_H = 120, 46
        n_layers = max(layer.values(), default=0) + 1
        layer_groups: dict[int, list[int]] = {}
        for t in result.tasks:
            l = layer.get(t.id, 0)
            layer_groups.setdefault(l, []).append(t.id)

        col_w = max(BOX_W + 50, (cw - 20) / max(n_layers, 1))
        positions: dict[int, tuple[float, float]] = {}
        assignee_colors: dict[str, str] = {}
        ci = 0

        for l_idx in range(n_layers):
            group = sorted(layer_groups.get(l_idx, []))
            col_h = len(group) * (BOX_H + 14) - 14
            start_y = max(10, (ch - col_h) / 2)
            cx = 10 + col_w * l_idx + BOX_W / 2
            for i, tid in enumerate(group):
                cy = start_y + i * (BOX_H + 14) + BOX_H / 2
                positions[tid] = (cx, cy)
                a = tasks_by_id[tid].assigned_to or "?"
                if a not in assignee_colors:
                    assignee_colors[a] = NODE_COLORS[ci % len(NODE_COLORS)]
                    ci += 1

        # edges
        for t in result.tasks:
            x2, y2 = positions.get(t.id, (0, 0))
            for dep_id in t.dependencies:
                if dep_id in positions:
                    x1, y1 = positions[dep_id]
                    self._canvas_arrow(canvas, x1 + BOX_W/2, y1, x2 - BOX_W/2, y2)

        # nodes
        for t in result.tasks:
            if t.id not in positions:
                continue
            cx, cy = positions[t.id]
            x1, y1 = cx - BOX_W/2, cy - BOX_H/2
            x2, y2 = cx + BOX_W/2, cy + BOX_H/2
            fill = assignee_colors.get(t.assigned_to or "?", "#444")
            r = 7
            pts = [x1+r,y1, x2-r,y1, x2,y1, x2,y1+r, x2,y2-r, x2,y2, x2-r,y2, x1+r,y2, x1,y2, x1,y2-r, x1,y1+r, x1,y1]
            canvas.create_polygon(pts, smooth=True, fill=fill, outline="#ccc", width=1.5)
            canvas.create_text(cx, cy - 10, text=f"#{t.id}", fill="white", font=("Consolas", 9, "bold"))
            desc = t.description[:16] + ("…" if len(t.description) > 16 else "")
            canvas.create_text(cx, cy + 3, text=desc, fill="#e0e0e0", font=("Consolas", 8))
            canvas.create_text(cx, cy + 16, text=t.assigned_to or "?", fill="#aaa", font=("Consolas", 7))

        # legend
        lx, ly = 6, ch - 6 - len(assignee_colors) * 15
        canvas.create_rectangle(lx-2, ly-3, lx+110, ly+len(assignee_colors)*15+1, fill="#111827", outline="#333")
        for name, col in assignee_colors.items():
            canvas.create_rectangle(lx, ly, lx+11, ly+11, fill=col, outline="")
            canvas.create_text(lx+15, ly+6, anchor="w", text=name, fill="#ccc", font=("Consolas", 8))
            ly += 15

    # ── Scheduler ───────────────────────────────────────────────────────────

    def _toggle_theme(self):
        current = ctk.get_appearance_mode()
        if current == "Dark":
            ctk.set_appearance_mode("light")
            self._theme_btn.configure(text="🌙 Dark")
        else:
            ctk.set_appearance_mode("dark")
            self._theme_btn.configure(text="☀ Light")

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select YAML Input",
            filetypes=[("YAML files", "*.yaml *.yml"), ("All files", "*.*")],
        )
        if path:
            self._yaml_path = path
            self._file_label.configure(text=os.path.basename(path), text_color=SUCCESS)
            self._log(f"Loaded: {path}")
            self._preview_yaml(path)
            self._load_yaml_into_editor(path)

    def _preview_yaml(self, path: str):
        try:
            data = load_yaml(path)
            n_tasks = len(data["tasks"])
            n_members = len(data["members"])
            self._summary_label.configure(
                text=(
                    f"Project: {data['start'].strftime('%b %d, %Y')} → "
                    f"{data['end'].strftime('%b %d, %Y')}\n"
                    f"Tasks: {n_tasks}   |   Members: {n_members}\n"
                    f"Max concurrent tasks/person: {data['max_tasks']}"
                )
            )
        except Exception as e:
            self._summary_label.configure(text=f"Error reading file:\n{e}")

    def _load_yaml_into_editor(self, path: str):
        """Load YAML into the Task Editor so tasks can be visually edited."""
        try:
            with open(path, "r") as f:
                raw = yaml.safe_load(f)

            import scheduler as sched
            self._project_start = sched._parse_date(str(raw.get("start_date", dt.date.today().strftime("%m/%d/%Y"))))
            self._project_end   = sched._parse_date(str(raw.get("end_date",   (dt.date.today() + dt.timedelta(days=30)).strftime("%m/%d/%Y"))))
            self._max_tasks     = int(raw.get("maximum_tasks", 2))
            self._members_raw   = raw.get("members", [])

            self._editor_tasks = []
            for t in raw.get("tasks", []):
                self._editor_tasks.append({
                    "id":           int(t["ID"]),
                    "description":  t.get("description", ""),
                    "areas":        [a.strip() for a in t.get("areas", [])],
                    "dependencies": [int(d) for d in t.get("dependencies", [])],
                    "estimated_days": int(t.get("estimated_time_days", 1)),
                })

            self._clear_task_form()
            self._refresh_task_editor()
            self._log(f"Editor loaded {len(self._editor_tasks)} tasks from {os.path.basename(path)}")
        except Exception as e:
            self._log(f"Editor load error: {e}")

    def _log(self, msg: str):
        ts = dt.datetime.now().strftime("%H:%M:%S")
        self._log_text.insert("end", f"[{ts}] {msg}\n")
        self._log_text.see("end")

    def _set_status(self, text: str, color: str = "gray"):
        self._status_label.configure(text=text, text_color=color)

    def _run_scheduler(self):
        if not self._yaml_path:
            messagebox.showwarning("No File", "Please select or save a YAML input file first.")
            return
        self._run_btn.configure(state="disabled")
        self._progress.set(0)
        self._set_status("Scheduling…", ACCENT)
        threading.Thread(target=self._scheduler_thread, daemon=True).start()

    def _scheduler_thread(self):
        try:
            self.after(0, lambda: self._progress.set(0.2))
            self._log("Parsing YAML…")
            data = load_yaml(self._yaml_path)

            self.after(0, lambda: self._progress.set(0.4))
            self._log("Building dependency graph…")
            G = build_graph(data["tasks"])
            topo = topological_order(G)
            self._log(f"Topological order: {topo}")

            self.after(0, lambda: self._progress.set(0.6))
            self._log("Assigning tasks to team members…")
            result = schedule(data)
            self._result = result

            self.after(0, lambda: self._progress.set(0.9))
            self._log(f"Scheduling complete. {len(result.warnings)} warning(s).")
            for w in result.warnings:
                self._log(f"  ⚠ {w}")

            self.after(0, lambda: self._display_results(result))
            self.after(0, lambda: self._progress.set(1.0))
            self.after(0, lambda: self._set_status("✓ Schedule ready", SUCCESS))
            self.after(0, lambda: self._export_btn.configure(state="normal"))
            self.after(0, lambda: self._cal_btn.configure(state="normal"))
            self.after(0, lambda: self._cal_quick_export.configure(state="normal"))
            self.after(0, lambda: self._run_btn.configure(state="normal"))

        except Exception as e:
            self.after(0, lambda: self._set_status(f"Error: {e}", ERROR))
            self._log(f"ERROR: {e}")
            self.after(0, lambda: self._run_btn.configure(state="normal"))
            self.after(0, lambda: self._progress.set(0))

    def _display_results(self, result: ScheduleResult):
        # Schedule table
        self._schedule_text.delete("1.0", "end")
        header = (
            f"{'ID':>4}  {'Description':<32} {'Assigned':<10} "
            f"{'Start':<12} {'End':<12} {'Days':>5}\n"
        )
        self._schedule_text.insert("end", header)
        self._schedule_text.insert("end", "─" * 85 + "\n")
        for t in sorted(result.tasks, key=lambda t: (t.start_date or dt.date.max)):
            line = (
                f"{t.id:>4}  {t.description:<32} {(t.assigned_to or '?'):<10} "
                f"{t.start_date.strftime('%m/%d/%Y') if t.start_date else 'N/A':<12} "
                f"{t.end_date.strftime('%m/%d/%Y') if t.end_date else 'N/A':<12} "
                f"{t.estimated_days:>5}\n"
            )
            self._schedule_text.insert("end", line)

        # ASCII Gantt
        if any(t.start_date for t in result.tasks):
            self._schedule_text.insert("end", "\n\n── Mini Gantt ─────────────────────────────────────────\n\n")
            min_d = min(t.start_date for t in result.tasks if t.start_date)
            max_d = max(t.end_date   for t in result.tasks if t.end_date)
            total = (max_d - min_d).days + 1
            scale = min(60, total)
            for t in sorted(result.tasks, key=lambda t: (t.start_date or dt.date.max)):
                if t.start_date and t.end_date:
                    s = int((t.start_date - min_d).days / total * scale)
                    e = int((t.end_date   - min_d).days / total * scale) + 1
                    bar = " " * s + "█" * (e - s)
                    label = f"[{t.id}] {t.description[:20]}"
                    self._schedule_text.insert("end", f"  {label:<26} |{bar}\n")

        # Dependency graph text
        self._graph_text.delete("1.0", "end")
        self._graph_text.insert("end", "Dependency Graph (Adjacency List)\n")
        self._graph_text.insert("end", "══════════════════════════════════\n\n")
        for tid in result.topo_order:
            task  = next(t for t in result.tasks if t.id == tid)
            preds = list(result.graph.predecessors(tid))
            succs = list(result.graph.successors(tid))
            self._graph_text.insert(
                "end",
                f"  Task {tid}: {task.description}\n"
                f"     ← depends on: {preds if preds else 'none'}\n"
                f"     → feeds into: {succs if succs else 'none'}\n\n"
            )
        self._graph_text.insert("end", "\nTopological Order: " +
                                " → ".join(str(t) for t in result.topo_order) + "\n")
        if nx.is_directed_acyclic_graph(result.graph):
            self._graph_text.insert("end", "\n✓ No circular dependencies detected.\n")

        # Draw visual graph
        self.after(200, lambda: self._draw_graph_canvas(result))

        # Calendar links
        self._cal_text.delete("1.0", "end")
        self._cal_text.insert("end", "Calendar Integration\n")
        self._cal_text.insert("end", "══════════════════════════════════════════════════════════\n\n")
        self._cal_text.insert("end", "  Download the .ics file to import into any calendar app:\n")
        self._cal_text.insert("end", "  Apple Calendar, Google Calendar, Outlook, Thunderbird, etc.\n\n")
        self._cal_text.insert("end", "  Use the '📅 Share to Calendar' button or 'Quick Export' above.\n\n")
        self._cal_text.insert("end", "─" * 60 + "\n\n")
        links = generate_calendar_links(result)
        for link in links:
            self._cal_text.insert("end", f"  📌 [Task {link['task_id']}] {link['description']}\n")
            self._cal_text.insert("end", f"     Assigned: {link['assigned_to']}  |  ")
            self._cal_text.insert("end",
                f"{link['start'].strftime('%b %d')} – {link['end'].strftime('%b %d, %Y')}\n")
            self._cal_text.insert("end", f"\n     Google Calendar:\n     {link['google_url']}\n\n")
            self._cal_text.insert("end", f"     Outlook Web:\n     {link['outlook_url']}\n\n")
            self._cal_text.insert("end", "  " + "─" * 56 + "\n\n")
        self._cal_text.insert("end", f"  Total events: {len(links)}\n")

    def _export(self):
        if not self._result:
            return
        fmt = self._format_var.get()
        ext_map  = {"xlsx": ".xlsx", "csv": ".csv", "ics": ".ics"}
        ext = ext_map.get(fmt, ".xlsx")
        ftype_map = {
            "xlsx": [("Excel files", "*.xlsx")],
            "csv":  [("CSV files",   "*.csv")],
            "ics":  [("iCalendar",   "*.ics")],
        }
        path = filedialog.asksaveasfilename(
            title="Save Gantt Chart" if fmt != "ics" else "Save Calendar File",
            defaultextension=ext,
            filetypes=ftype_map.get(fmt, []) + [("All files", "*.*")],
            initialfile=f"gantt_chart{ext}" if fmt != "ics" else "taskflow_schedule.ics",
        )
        if not path:
            return
        try:
            if fmt == "xlsx":   export_gantt_xlsx(self._result, path)
            elif fmt == "ics":  export_ics(self._result, path)
            else:               export_gantt_csv(self._result, path)
            self._log(f"Exported to {path}")
            self._set_status(f"✓ Exported: {os.path.basename(path)}", SUCCESS)
            msg = f"Saved to:\n{path}"
            if fmt == "ics":
                msg += "\n\nTo import:\n• Double-click the .ics file, or\n• Drag into your calendar app."
            messagebox.showinfo("Export Complete", msg)
        except Exception as e:
            self._log(f"Export error: {e}")
            messagebox.showerror("Export Error", str(e))

    def _quick_export_ics(self):
        if not self._result:
            return
        path = filedialog.asksaveasfilename(
            title="Save Calendar File (.ics)",
            defaultextension=".ics",
            filetypes=[("iCalendar files", "*.ics"), ("All files", "*.*")],
            initialfile="taskflow_schedule.ics",
        )
        if not path:
            return
        try:
            export_ics(self._result, path)
            self._log(f"ICS exported to {path}")
            self._set_status(f"✓ Calendar saved: {os.path.basename(path)}", SUCCESS)
            messagebox.showinfo("Calendar Exported",
                f"Calendar file saved to:\n{path}\n\nDouble-click to open in your default calendar app.")
        except Exception as e:
            self._log(f"ICS export error: {e}")
            messagebox.showerror("Export Error", str(e))

    def _open_calendar_dialog(self):
        if not self._result:
            return
        dialog = ctk.CTkToplevel(self)
        dialog.title("Share to Calendar")
        dialog.geometry("480x420")
        dialog.transient(self)
        dialog.grab_set()

        header = ctk.CTkFrame(dialog, height=50, corner_radius=0, fg_color="#7C3AED")
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header, text="📅  Share to Calendar",
            font=ctk.CTkFont(size=16, weight="bold"), text_color="white",
        ).pack(side="left", padx=16)

        body = ctk.CTkFrame(dialog, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=16)

        card1 = ctk.CTkFrame(body, corner_radius=10)
        card1.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(card1, text="Download .ics File",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(padx=16, pady=(12, 2), anchor="w")
        ctk.CTkLabel(
            card1,
            text="Works with Apple Calendar, Google Calendar, Outlook,\nThunderbird, and any iCalendar-compatible app.",
            font=ctk.CTkFont(size=11), text_color="gray", justify="left",
        ).pack(padx=16, pady=(0, 4), anchor="w")

        ics_row = ctk.CTkFrame(card1, fg_color="transparent")
        ics_row.pack(fill="x", padx=16, pady=(4, 12))
        ctk.CTkButton(
            ics_row, text="⬇  Save .ics File", height=36,
            fg_color="#7C3AED", hover_color="#6D28D9",
            command=lambda: [dialog.destroy(), self._quick_export_ics()],
        ).pack(side="left")
        ctk.CTkLabel(ics_row, text="Reminder:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(20, 6))
        self._alarm_var = ctk.StringVar(value="30")
        ctk.CTkOptionMenu(ics_row, values=["None", "15", "30", "60", "120", "1440"],
                          variable=self._alarm_var, width=80, height=30).pack(side="left")
        ctk.CTkLabel(ics_row, text="min", font=ctk.CTkFont(size=11), text_color="gray").pack(side="left", padx=(4, 0))

        card2 = ctk.CTkFrame(body, corner_radius=10)
        card2.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(card2, text="Web Calendar Links",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(padx=16, pady=(12, 2), anchor="w")
        ctk.CTkLabel(
            card2,
            text="Open individual tasks in Google Calendar or Outlook Web.\nLinks are shown in the Calendar tab.",
            font=ctk.CTkFont(size=11), text_color="gray", justify="left",
        ).pack(padx=16, pady=(0, 4), anchor="w")

        links_row = ctk.CTkFrame(card2, fg_color="transparent")
        links_row.pack(fill="x", padx=16, pady=(4, 12))

        def _copy_all_links():
            links = generate_calendar_links(self._result)
            text_parts = [
                f"[Task {lnk['task_id']}] {lnk['description']}\n"
                f"  Google: {lnk['google_url']}\n"
                f"  Outlook: {lnk['outlook_url']}\n"
                for lnk in links
            ]
            self.clipboard_clear()
            self.clipboard_append("\n".join(text_parts))
            self._log("Calendar links copied to clipboard")
            self._set_status("✓ Links copied to clipboard", SUCCESS)
            dialog.destroy()

        ctk.CTkButton(links_row, text="📋  Copy All Links", height=36,
                      fg_color=ACCENT, hover_color="#3d6a96",
                      command=_copy_all_links).pack(side="left", padx=(0, 8))
        ctk.CTkButton(links_row, text="↗  Go to Calendar Tab", height=36,
                      fg_color="transparent", border_width=1, border_color=ACCENT,
                      text_color=ACCENT, hover_color="#1a2540",
                      command=lambda: [dialog.destroy(), self._tabview.set("📅 Calendar")],
                      ).pack(side="left")


def main():
    app = TaskFlowApp()
    app.mainloop()


if __name__ == "__main__":
    main()
