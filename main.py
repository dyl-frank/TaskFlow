#!/usr/bin/env python3
"""
TaskFlow – Project Scheduler
─────────────────────────────
Usage:
    python main.py                        # Launch web GUI at http://127.0.0.1:5000
    python main.py --port 8080            # Launch on custom port
    python main.py input.yaml             # CLI: schedule & export xlsx
    python main.py input.yaml -o out.csv  # CLI: export to CSV
    python main.py input.yaml --fmt ics   # CLI: export to ICS
"""

import argparse, sys, os


def cli_run(args):
    from scheduler import load_yaml, schedule, export_gantt_xlsx, export_gantt_csv, export_ics

    print(f"TaskFlow Scheduler")
    print(f"{'─' * 40}")
    print(f"Input:  {args.input}")

    data = load_yaml(args.input)
    print(f"Project: {data['start']} → {data['end']}")
    print(f"Tasks: {len(data['tasks'])}  |  Members: {len(data['members'])}")

    result = schedule(data)
    print(f"\nTopological order: {' → '.join(str(t) for t in result.topo_order)}")

    print(f"\n{'ID':>4}  {'Description':<32} {'Assigned':<10} {'Start':<12} {'End':<12} {'Days':>5}")
    print("─" * 85)
    for t in sorted(result.tasks, key=lambda t: (t.start_date or __import__('datetime').date.max)):
        sd = t.start_date.strftime('%m/%d/%Y') if t.start_date else 'N/A'
        ed = t.end_date.strftime('%m/%d/%Y') if t.end_date else 'N/A'
        print(f"{t.id:>4}  {t.description:<32} {(t.assigned_to or '?'):<10} {sd:<12} {ed:<12} {t.estimated_days:>5}")

    if result.warnings:
        print(f"\n⚠  Warnings:")
        for w in result.warnings:
            print(f"   • {w}")

    fmt = args.fmt or ("csv" if args.output and args.output.endswith(".csv")
                       else "ics" if args.output and args.output.endswith(".ics")
                       else "xlsx")
    output = args.output or f"gantt_chart.{fmt}"

    if fmt == "csv":
        export_gantt_csv(result, output)
    elif fmt == "ics":
        export_ics(result, output)
        print(f"\n✓ Calendar file exported to: {output}")
        print(f"  Import by double-clicking the .ics file or dragging into your calendar app.")
        return
    else:
        export_gantt_xlsx(result, output)
    print(f"\n✓ Gantt chart exported to: {output}")


def main():
    # If first arg is a file path (not a flag), run in CLI mode
    if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
        parser = argparse.ArgumentParser(description="TaskFlow Project Scheduler")
        parser.add_argument("input", help="Path to YAML input file")
        parser.add_argument("-o", "--output", help="Output file path")
        parser.add_argument("--fmt", choices=["xlsx", "csv", "ics"], help="Output format (xlsx, csv, or ics)")
        args = parser.parse_args()
        cli_run(args)
    else:
        # Web GUI mode
        parser = argparse.ArgumentParser(description="TaskFlow Project Scheduler – Web GUI")
        parser.add_argument("--port", type=int, default=5000, help="Port to run the web server on (default: 5000)")
        parser.add_argument("--debug", action="store_true", help="Run in Flask debug mode")
        args = parser.parse_args()

        from webapp import run_webapp
        run_webapp(port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
