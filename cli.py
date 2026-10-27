"""
Command-line entry point (no GUI) — runs the same pipeline as main.py.

    python cli.py sample_data/sample_inventory_report.xlsx [--out OUTPUT_FOLDER]
"""
import argparse

import excel_writer
import processor


def run(input_path, output_folder=None):
    df = processor.read_asl_file(input_path)
    column_map, missing, distance_column = processor.resolve_columns(df)
    if missing:
        raise SystemExit(f'Required column "{missing[0]}" not found.')

    pg_df = processor.filter_pg_rows(df, column_map["Type Trim"])
    print(f"PG rows: {len(pg_df):,}")
    detail_df = processor.generate_detail_report(pg_df, column_map, distance_column)
    if detail_df.empty:
        print("No optimization issues found.")
        return
    summary_df = processor.generate_summary_report(detail_df)
    output_path = excel_writer.build_output_path(input_path, output_folder)
    excel_writer.write_output_excel(summary_df, detail_df, output_path)

    print(f"Rows needing attention: {len(detail_df):,}")
    print(detail_df["Issue Type"].str.split("; ").explode().value_counts().to_string())
    print(f"Report saved to: {output_path}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()
    run(args.input, args.out)
