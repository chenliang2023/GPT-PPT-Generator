from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from pptx import Presentation
from pptx.exc import PackageNotFoundError

SCRIPT_DIR = Path(__file__).resolve().parent
EDITABLE_PPTX_SCRIPTS = SCRIPT_DIR.parents[1] / "images-to-editable-pptx" / "scripts"
sys.path.insert(0, str(EDITABLE_PPTX_SCRIPTS))

from build_editable_pptx import build_pptx  # noqa: E402


def _template_with_blank_layout(presentation: Presentation) -> Presentation:
    layouts = list(presentation.slide_layouts)
    blank_layout = next(
        (layout for layout in layouts if "blank" in layout.name.lower() or "空白" in layout.name.lower()),
        None,
    )
    if blank_layout is None:
        if len(layouts) < 7:
            raise ValueError(
                "Template does not contain a blank layout and has fewer than 7 layouts"
            )
        blank_layout = layouts[6]

    class LayoutsProxy:
        def __getitem__(self, index):
            if index == 6:
                return blank_layout
            return layouts[index]

        def __iter__(self):
            return iter(layouts)

        def __len__(self):
            return len(layouts)

    class TemplatePresentationProxy:
        def __init__(self, original: Presentation) -> None:
            self._original = original
            self.slide_layouts = LayoutsProxy()

        def __getattr__(self, name):
            return getattr(self._original, name)

    return TemplatePresentationProxy(presentation)


def _warn_if_slide_size_differs(spec: dict, presentation: Presentation) -> None:
    slide_size = spec.get("slide_size")
    if not isinstance(slide_size, dict):
        return

    width = slide_size.get("width")
    height = slide_size.get("height")
    if width is None and height is None:
        return

    template_width = presentation.slide_width / 914400
    template_height = presentation.slide_height / 914400
    spec_width = float(width) if width is not None else template_width
    spec_height = float(height) if height is not None else template_height
    if spec_width != template_width or spec_height != template_height:
        print(
            "Warning: template slide size takes precedence over spec slide_size",
            file=sys.stderr,
        )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build an editable PPTX from a native reconstruction spec."
    )
    parser.add_argument("--spec", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--template", type=Path)
    parser.add_argument("--allow-full-bleed-images", action="store_true")
    args = parser.parse_args()

    try:
        spec = json.loads(args.spec.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"JSON parse failed: {exc}", file=sys.stderr)
        return 1
    except OSError as exc:
        print(f"Unable to read spec: {exc}", file=sys.stderr)
        return 1

    presentation = None
    if args.template is not None:
        try:
            presentation = Presentation(args.template)
        except (OSError, PackageNotFoundError, ValueError, KeyError, TypeError) as exc:
            print(f"Unable to read template {args.template}: {exc}", file=sys.stderr)
            return 1
        try:
            _warn_if_slide_size_differs(spec, presentation)
            build_pptx(
                spec=spec,
                output_path=args.out,
                base_dir=args.spec.parent,
                allow_full_bleed_images=args.allow_full_bleed_images,
                existing_presentation=_template_with_blank_layout(presentation),
                preserve_slide_size=True,
            )
        except (OSError, ValueError, IndexError, TypeError) as exc:
            print(f"Unable to build PPTX from template: {exc}", file=sys.stderr)
            return 1
    else:
        try:
            build_pptx(
                spec=spec,
                output_path=args.out,
                base_dir=args.spec.parent,
                allow_full_bleed_images=args.allow_full_bleed_images,
            )
        except (OSError, ValueError, IndexError, TypeError) as exc:
            print(f"Unable to build PPTX: {exc}", file=sys.stderr)
            return 1

    print(f"Created editable PPTX: {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
