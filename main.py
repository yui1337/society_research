from argparse import ArgumentParser, Namespace

from handle_xlsx import parse_xlsx, create_result_table, save_xlsx


def parse_args() -> Namespace:
    parser = ArgumentParser()

    parser.add_argument("input_xlsx")
    parser.add_argument("output_xlsx")
    parser.add_argument("group", type=int)

    return parser.parse_args()


def main():
    args = parse_args()

    table = parse_xlsx(args.input_xlsx)
    result_table = create_result_table(table, args.group)
    save_xlsx(result_table, args.output_xlsx)


if __name__ == "__main__":
    main()