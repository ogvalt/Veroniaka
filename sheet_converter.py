import argparse
import pandas as pd


def get_single_room(df_data, ids):
    df_room = df_data.T[[0, 1, ids + 2]]
    room_name = df_room[ids + 2][0]
    return df_room, room_name


def main(input_filename, input_sheet, number_of_header, output_file) -> None:
    INPUT_FILENAME = input_filename
    INPUT_SHEET: str = input_sheet
    NUMBER_OF_HEADERS: int = number_of_header
    OUTPUT_FILENAME = output_file

    df = pd.read_excel(INPUT_FILENAME, sheet_name=INPUT_SHEET, header=None)

    end_ids = df.T.shape[1] - NUMBER_OF_HEADERS

    print(f"Expected to process rooms: {end_ids}")

    filename = OUTPUT_FILENAME

    print("Processing...")

    df_room, room_name = get_single_room(df, ids=0)

    df_room.to_excel(filename, sheet_name=room_name, header=False, index=False)

    with pd.ExcelWriter(filename, mode="a", if_sheet_exists="replace") as writer:
        for idx in range(1, end_ids):
            df_room, room_name = get_single_room(df, idx)
            df_room.to_excel(writer, sheet_name=room_name, header=False, index=False)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert sheet into sheets')
    parser.add_argument('--input-filename', required=True,
                        type=str, help='Input filename to convert')
    parser.add_argument('--input-sheet', default="Sheet1",
                        type=str, help='Sheet to take data from')
    parser.add_argument('--number-of-header', default=2,
                        type=int, help='Number of headers in input data')

    parser.add_argument('--output-file', default="results.xlsx",
                        type=str, help='Results file')
    args = parser.parse_args()

    input_filename, input_sheet = args.input_filename, args.input_sheet
    number_of_header, output_file = args.number_of_header, args.output_file

    print("Starting")
    main(input_filename, input_sheet, number_of_header, output_file)
    print("Done!")
