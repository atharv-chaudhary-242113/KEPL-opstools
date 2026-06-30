import argparse
import sys

from bom_dic.bom_merge.main_merge import run_consolidator
from bom_dic.bom_split.main_split import run_classifier


def interactive_mode() -> None:
    """Interactive terminal wizard for non-technical users."""
    print("=== Welcome to KEPL OpsTools ===")

    while True:
        tool_choice = (
            input("What tool would you like to use? (Options: 'bom' or 'quit'): ")
            .strip()
            .lower()
        )

        if tool_choice == "bom":
            action = (
                input("Do you want to 'split' or 'merge'? (or type 'back'): ")
                .strip()
                .lower()
            )

            if action in ["s", "split"]:
                sub = (
                    input(
                        "Do you want to split sub-categories (e.g., -1, -2, -A)? (y/n): "  # noqa: E501
                    )
                    .strip()
                    .lower()
                )
                split_sub = sub in ["y", "yes"]
                print(
                    "\n[*] Proceeding with BoM Split using default input directory..."
                )
                run_classifier("DEFAULT", split_sub)
                return
            elif action in ["m", "merge"]:
                print(
                    "\n[*] Proceeding with BoM Merge using default input directory..."
                )
                run_consolidator("DEFAULT")
                return
            elif action == "back":
                break
            else:
                print("[-] Invalid choice. Please type 'split', 'merge', or 'back'.")

        elif tool_choice in ["quit", "q", "exit"]:
            print("Exiting OpsTools.")
            sys.exit(0)
        else:
            print(
                f"[-] Tool '{tool_choice}' not recognized. Currently available: 'bom'"
            )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="KEPL OpsTools Central Command",
        epilog="Run without arguments to enter the interactive terminal menu.",
    )

    parser.add_argument(
        "--split",
        nargs="?",
        const="DEFAULT",
        default=None,
        metavar="PATH",
        help="Custom path to the master BoM file/folder to split.",
    )

    parser.add_argument(
        "--split-sub",
        action="store_true",
        help="Force splitting of sub-categories (e.g. -1, -2, -A) during split operation.",  # noqa: E501
    )

    parser.add_argument(
        "--merge",
        nargs="?",
        const="DEFAULT",
        default=None,
        metavar="PATH",
        help="Custom path to the directory/file to consolidate.",
    )

    args = parser.parse_args()

    # FIXED: Now checks if split, split_sub, OR merge were triggered
    if args.split is not None or args.split_sub or args.merge is not None:
        # If split-sub is true but --split was omitted, fall back to DEFAULT path
        if args.split is not None or args.split_sub:
            target_path = args.split if args.split is not None else "DEFAULT"
            print(
                f"[*] CLI Triggered: BoM Split {'(Including Sub-categories)' if args.split_sub else ''}"  # noqa: E501
            )
            run_classifier(target_path, args.split_sub)

        if args.merge is not None:
            print("[*] CLI Triggered: BoM Merge")
            run_consolidator(args.merge)
    else:
        interactive_mode()


if __name__ == "__main__":
    main()
