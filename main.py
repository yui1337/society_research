from argparse import ArgumentParser, Namespace

from handle_xlsx import parse_xlsx, create_result_table, save_xlsx


def parse_args() -> Namespace:
    parser = ArgumentParser()

    parser.add_argument("input_xlsx", help="Path to input file")
    parser.add_argument("output_xlsx", help="Path to output file")
    parser.add_argument("group_size", type=int, help="Group size")
    parser.add_argument("group_id", type=int, help="Grou ID")

    return parser.parse_args()


def main():
    args = parse_args()

    table = parse_xlsx(args.input_xlsx)
    result_table = create_result_table(table, args.group_size, args.group_id)
    save_xlsx(result_table, args.output_xlsx)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {str(e)}")

