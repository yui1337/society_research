from argparse import ArgumentParser, Namespace

from handle_xlsx import parse_xlsx, create_result_table, save_xlsx


def parse_args() -> Namespace:
    parser = ArgumentParser()

    parser.add_argument("input_xlsx", type=str, help="Path to input file")
    parser.add_argument("output_xlsx", type=str, help="Path to output file")
    parser.add_argument("group_size",type=int, help="Group size")
    parser.add_argument("group_id", type=str, help="Group ID")

    return parser.parse_args()

def do_research(input_xlsx: str, output_xlsx: str, group_id: str, group_size: int):
    table = parse_xlsx(input_xlsx)
    result_table = create_result_table(table, group_size, group_id)
    save_xlsx(result_table, output_xlsx)

def main():
    args = parse_args()
    do_research(args.input_xlsx, args.output_xlsx, args.group_id, args.group_size)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {str(e)}")

