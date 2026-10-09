import argparse

from .renderer import render_presentation


def main(argv=None):
    p = argparse.ArgumentParser(description="Render pptx slides as png images")
    p.add_argument("pptx")
    p.add_argument("output_dir")
    p.add_argument("--width", type=int, default=1920, help="image width in pixels")
    args = p.parse_args(argv)
    for path in render_presentation(args.pptx, args.output_dir, width=args.width):
        print(path)


if __name__ == "__main__":
    main()
