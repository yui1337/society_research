from argparse import ArgumentParser, Namespace

from handle_xlsx import parse_xlsx, create_result_table, save_xlsx


def parse_args() -> Namespace:
    parser = ArgumentParser()

    parser.add_argument("input_xlsx")
    parser.add_argument("output_xlsx")
    parser.add_argument("group", type=int)

    return parser.parse_args()


def main():
    #args = parse_args()

   # table = parse_xlsx(args.input_xlsx)
    #result_table = create_result_table(table, args.group)
    #save_xlsx(result_table, args.output_xlsx)
    input_table = parse_xlsx(r"C:\Torrent\society_research\google_doc_raw.xlsx")
    create_result_table(input_table, 25)


if __name__ == "__main__":
    main()