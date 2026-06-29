from pathlib import Path
import argparse

from rocket_mass_properties import calculate, write_output, write_plots


def main() -> None:
    parser = argparse.ArgumentParser(description="QJ 1080A liquid rocket mass, CG and inertia calculator")
    parser.add_argument("--input", default="input_template.xlsx", help="input Excel workbook")
    parser.add_argument("--output", default=str(Path("output") / "mass_properties_result.xlsx"), help="output Excel workbook")
    parser.add_argument("--plots", default=str(Path("output") / "mass_properties_plots.pdf"), help="output PDF plots")
    parser.add_argument("--no-plots", action="store_true", help="skip PDF plot generation")
    args = parser.parse_args()

    results = calculate(args.input)
    write_output(results, args.output)
    if not args.no_plots:
        write_plots(results, args.plots)
    print(f"saved {Path(args.output).resolve()}")
    if not args.no_plots:
        print(f"saved {Path(args.plots).resolve()}")


if __name__ == "__main__":
    main()
