import sys
import argparse
from uln_compiler import ULNCompiler

def main():
    parser = argparse.ArgumentParser(description="ULN to DOCX Formatter CLI (pywin32)")
    parser.add_argument("input", help="Path to input ULN text file (.txt or .uln)")
    parser.add_argument("output", help="Path to output Word document (.docx)")
    parser.add_argument("--pdf", action="store_true", help="Also export as PDF")
    parser.add_argument("--font", default="Times New Roman", help="Font family name")
    parser.add_argument("--size", type=float, default=12.0, help="Font size in pt")
    parser.add_argument("--margin-top", type=float, default=2.0, help="Top margin in cm")
    parser.add_argument("--margin-bottom", type=float, default=2.0, help="Bottom margin in cm")
    parser.add_argument("--margin-left", type=float, default=3.0, help="Left margin in cm")
    parser.add_argument("--margin-right", type=float, default=1.5, help="Right margin in cm")

    args = parser.parse_args()

    settings = {
        "font_name": args.font,
        "font_size": args.size,
        "margin_top": args.margin_top,
        "margin_bottom": args.margin_bottom,
        "margin_left": args.margin_left,
        "margin_right": args.margin_right,
    }

    compiler = ULNCompiler(settings)
    out_docx = compiler.compile_file(args.input, args.output)
    print(f"[Success] Compiled DOCX saved to: {out_docx}")

    if args.pdf:
        pdf_path = args.output.rsplit(".", 1)[0] + ".pdf"
        ok = compiler.convert_docx_to_pdf(out_docx, pdf_path)
        if ok:
            print(f"[Success] Exported PDF saved to: {pdf_path}")
        else:
            print("[Error] Failed to export PDF.")

if __name__ == "__main__":
    main()
