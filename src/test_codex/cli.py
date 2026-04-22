"""Command-line interface for the Excel note combiner."""

from __future__ import annotations

import argparse
import os

from test_codex.chatgpt_excel import run_chatgpt_prompt_from_workbook
from test_codex.excel_agent import NoteMapping, apply_note_mappings, load_mappings


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Combine notes from specific Excel cells into another cell."
    )
    parser.add_argument(
        "source_workbook",
        help="Path to the source .xlsx workbook. In ChatGPT mode, this is the prompt workbook.",
    )
    parser.add_argument(
        "output_workbook",
        help="Path to the output .xlsx workbook. In ChatGPT mode, this receives the response.",
    )
    parser.add_argument(
        "--cells",
        nargs="+",
        help="Source cell addresses to read, such as A1 B1 C1.",
    )
    parser.add_argument(
        "--source-sheet",
        help="Worksheet name to read from. Defaults to the source workbook's active sheet.",
    )
    parser.add_argument(
        "--target-sheet",
        help="Worksheet name in the output workbook to write into.",
    )
    parser.add_argument(
        "--target-cell",
        help="Target cell address where the combined note text should be written.",
    )
    parser.add_argument(
        "--separator",
        default="\n\n",
        help="Text to place between notes. Defaults to a blank line.",
    )
    parser.add_argument(
        "--cell-values",
        action="store_true",
        help="Read visible cell values instead of right-click Excel notes.",
    )
    parser.add_argument(
        "--mapping",
        help="CSV file describing multiple note transfers to apply in one run.",
    )
    parser.add_argument(
        "--prompt-range",
        help="Cell range to read as a ChatGPT prompt, such as A1:D8.",
    )
    parser.add_argument(
        "--prompt-sheet",
        help="Worksheet name containing the prompt range. Defaults to the source active sheet.",
    )
    parser.add_argument(
        "--response-sheet",
        help="Worksheet name in the output workbook where the ChatGPT response should be written.",
    )
    parser.add_argument(
        "--response-cell",
        help="Cell in the output workbook where the ChatGPT response should be written.",
    )
    parser.add_argument(
        "--model",
        default=os.environ.get("OPENAI_MODEL", "gpt-4.1-mini"),
        help="OpenAI model to use. Defaults to OPENAI_MODEL or gpt-4.1-mini.",
    )
    parser.add_argument(
        "--instructions",
        help="Optional developer/system-style instructions for the OpenAI response.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)

    if args.prompt_range:
        if not args.response_sheet or not args.response_cell:
            raise SystemExit(
                "--response-sheet and --response-cell are required with --prompt-range"
            )
        prompt, response_text = run_chatgpt_prompt_from_workbook(
            args.source_workbook,
            args.output_workbook,
            prompt_range=args.prompt_range,
            prompt_sheet_name=args.prompt_sheet or args.source_sheet,
            response_sheet_name=args.response_sheet,
            response_cell=args.response_cell,
            model=args.model,
            instructions=args.instructions,
        )
        print(f"Sent prompt from {args.prompt_range}:")
        print(prompt)
        print(f"Wrote ChatGPT response to {args.response_sheet}!{args.response_cell}:")
        print(response_text)
        return

    if args.mapping:
        mappings = load_mappings(args.mapping)
    else:
        if not args.cells or not args.target_sheet or not args.target_cell:
            raise SystemExit(
                "--cells, --target-sheet, and --target-cell are required "
                "unless --mapping is provided"
            )
        mappings = [
            NoteMapping(
                source_cells=args.cells,
                source_sheet=args.source_sheet,
                target_sheet=args.target_sheet,
                target_cell=args.target_cell,
            )
        ]

    results = apply_note_mappings(
        args.source_workbook,
        args.output_workbook,
        mappings,
        default_source_sheet_name=args.source_sheet,
        separator=args.separator,
        use_cell_values=args.cell_values,
    )

    for mapping, combined_note in results:
        print(f"Wrote combined note to {mapping.target_sheet}!{mapping.target_cell}:")
        print(combined_note)


if __name__ == "__main__":
    main()
