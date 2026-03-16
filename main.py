
import argparse
import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from extractor import parse_inspection_pdf, parse_thermal_pdf
from analyzer import analyze_with_ai
from report_generator import generate_report


def print_banner():
    print("\n" + "═" * 60)
    print("   DDR Report Generator — UrbanRoof Inspection System")
    print("═" * 60)


def print_step(step, total, message):
    print(f"\n[{step}/{total}] {message}...")


def print_ok(message):
    print(f"    ✓ {message}")


def print_warn(message):
    print(f"    ⚠ {message}")


def main():
    parser = argparse.ArgumentParser(description="DDR Report Generator")
    parser.add_argument("--inspection", required=True, help="Path to Inspection Report PDF")
    parser.add_argument("--thermal",    required=True, help="Path to Thermal Images PDF")
    parser.add_argument("--output",     default="output/DDR_Report.docx", help="Output DOCX path")
    parser.add_argument("--api-key",    default=None,
                        help="OpenRouter key (sk-or-...) or Anthropic key (sk-ant-...)")
    parser.add_argument("--assets-dir", default="assets", help="Directory for extracted images")
    args = parser.parse_args()

    print_banner()

    # ── Validate inputs ──
    if not os.path.exists(args.inspection):
        print(f"ERROR: Inspection PDF not found: {args.inspection}"); sys.exit(1)
    if not os.path.exists(args.thermal):
        print(f"ERROR: Thermal PDF not found: {args.thermal}"); sys.exit(1)

    # ── API key auto-detection ──
    if args.api_key:
        if args.api_key.startswith("sk-or-"):
            os.environ["OPENROUTER_API_KEY"] = args.api_key
            print_ok("OpenRouter key detected (sk-or-...) → AI via openrouter.ai/api/v1")
        elif args.api_key.startswith("sk-ant-"):
            os.environ["ANTHROPIC_API_KEY"] = args.api_key
            print_ok("Anthropic key detected (sk-ant-...) → AI via api.anthropic.com")
        else:
            # Unknown prefix — try OpenRouter first
            os.environ["OPENROUTER_API_KEY"] = args.api_key
            print_ok("API key set → trying OpenRouter first")
    elif os.environ.get("OPENROUTER_API_KEY"):
        print_ok("OPENROUTER_API_KEY found in environment — AI enabled")
    elif os.environ.get("ANTHROPIC_API_KEY"):
        print_ok("ANTHROPIC_API_KEY found in environment — AI enabled")
    else:
        print_warn("No API key — using rule-based analysis")
        print_warn("Pass --api-key sk-or-... or set OPENROUTER_API_KEY to enable AI")

    total_steps = 5
    t_start = time.time()

    # ── Step 1: Parse Inspection PDF ──
    print_step(1, total_steps, "Parsing Inspection Report PDF")
    inspection_data = parse_inspection_pdf(args.inspection, args.assets_dir)
    print_ok(f"Inspection date:     {inspection_data.inspection_date}")
    print_ok(f"Inspected by:        {inspection_data.inspected_by}")
    print_ok(f"Property type:       {inspection_data.property_type}")
    print_ok(f"Overall score:       {inspection_data.overall_score}")
    print_ok(f"Impacted areas:      {len(inspection_data.impacted_areas)}")
    print_ok(f"Photos extracted:    {len(inspection_data.all_images)}")
    print_ok(f"Checklist items:     {len(inspection_data.checklist_items)}")

    # ── Step 2: Parse Thermal PDF ──
    print_step(2, total_steps, "Parsing Thermal Images PDF")
    thermal_data = parse_thermal_pdf(args.thermal, args.assets_dir)
    if thermal_data.readings:
        print_ok(f"Thermal readings:    {len(thermal_data.readings)}")
        print_ok(f"Max hotspot:         {thermal_data.max_hotspot} °C")
        print_ok(f"Min coldspot:        {thermal_data.min_coldspot} °C")
        print_ok(f"Avg delta:           {thermal_data.avg_delta} °C")
        print_ok(f"Anomalies (Δ≥4°C):  {thermal_data.anomaly_count}")
    else:
        print_warn("No thermal readings could be extracted")

    # ── Step 3: AI / Rule-based Analysis ──
    print_step(3, total_steps, "Analysing findings")
    analysis = analyze_with_ai(inspection_data, thermal_data)
    print_ok(f"Severity:            {analysis.get('severity_assessment','N/A')[:60]}")
    print_ok(f"Findings generated:  {len(analysis.get('findings', []))}")
    print_ok(f"Recommendations:     {len(analysis.get('recommendations', []))}")

    # ── Step 4: Validate data quality ──
    print_step(4, total_steps, "Validating data quality")
    warnings = []
    for area in inspection_data.impacted_areas:
        if not area.negative_photos:
            warnings.append(f"Area {area.area_number}: No negative-side photos found")
        if not area.positive_photos:
            warnings.append(f"Area {area.area_number}: No positive-side photos found")
    for i, reading in enumerate(thermal_data.readings):
        if reading.thermal_image_path == "Image Not Available":
            warnings.append(f"Thermal reading {i+1}: Thermal image not available")
        if reading.visual_image_path == "Image Not Available":
            warnings.append(f"Thermal reading {i+1}: Visual photo not available")

    if warnings:
        for w in warnings[:10]: print_warn(w)
        if len(warnings) > 10: print_warn(f"... and {len(warnings)-10} more warnings")
    else:
        print_ok("All data validated — no missing items")

    # ── Step 5: Generate Report ──
    print_step(5, total_steps, f"Generating DOCX report → {args.output}")
    output_path = generate_report(inspection_data, thermal_data, analysis, args.output)
    elapsed = round(time.time() - t_start, 1)
    size_kb = os.path.getsize(output_path) // 1024

    print_ok(f"Report saved:  {output_path}")
    print_ok(f"File size:     {size_kb} KB")
    print_ok(f"Time elapsed:  {elapsed}s")

    print("\n" + "═" * 60)
    print("   DDR Report generated successfully!")
    print(f"   Output: {os.path.abspath(output_path)}")
    print("═" * 60 + "\n")

    return output_path


if __name__ == "__main__":
    main()
