import datetime
import os
import pandas as pd


def get_current_iso_week():
    """
    Gets the current ISO week number and year in "WWYY'ZZ" format.

    Returns:
        str: The current week number and year in "WWYY'ZZ" format.
    """
    current_date = datetime.datetime.now()
    iso_year, iso_week, _ = current_date.isocalendar()
    week_str = f"WW{iso_week:02d}"
    year_str = str(iso_year)[-2:]
    return f"{week_str}'{year_str}"


def get_output_path(file_name):
    """
    Gets the path to the future ship mode file based on the current ISO week.

    Args:
        file_name (str): The name of the desired file.

    Returns:
        str: The full path to the future ship mode file.
    """
    current_week = get_current_iso_week()
    base_path = r"//sgsind0nsifsv01a/IMAC Data/IMAC Senior or Teams/Europe & Other Asia Team - SP/VN01/FUTURE SHIP MODE/"
    return os.path.join(base_path, current_week, f"JB {file_name}.xlsx")


def get_oh_part_path():
    # Get the current date
    now = datetime.datetime.now()
     # Calculate the Monday date of the current week
    monday = now - datetime.timedelta(days=now.weekday())
    # Format the Monday date as MMDD
    monday_str = monday.strftime("%m%d")
    expanded_path = r"//sgsind0nsifsv01a/IMAC Data/IMAC Senior or Teams/Europe & Other Asia Team - SP/VN01/VN01 on hand part"
    return os.path.join(expanded_path, f"VN01 OH {monday_str}.xlsx")


def get_future_shipmode_path(file_name):
    """
    Gets the path to the output file based on the current ISO week.

    Args:
        file_name (str): The name of the desired output file.

    Returns:
        str: The full path to the output file.
    """
    current_week = get_current_iso_week()
    base_path = r"//sgsind0nsifsv01a/IMAC Data/IMAC Senior or Teams/Europe & Other Asia Team - SP/VN01/FUTURE SHIP MODE/"
    return os.path.join(base_path, current_week, f"{file_name}.xlsx")


def main():
    file_name = input("Enter file name: ")

    full_file_path_sm = get_future_shipmode_path(file_name)
    full_file_path_oh = get_oh_part_path()

    df_oh = pd.read_excel(full_file_path_oh)
    df_oh.columns = df_oh.columns.str.replace("Material", "Customer Part.")

    df_sm = pd.read_excel(full_file_path_sm)
    df_sm["Customer Part."] = df_sm["Customer Part."].apply(str)

    merged_df = pd.merge(df_oh, df_sm, on="Customer Part.")
    filtered_df = merged_df[merged_df["PGr"].str.startswith("U")]

    # Define desired columns before dropping others
    desired_columns =  ["Ship To", "Order Number", "Line Number", "Customer Part.",
                       "Mfr", "Mfr Part.", "Qty Required", "Qty Allocated",
                       "Date Ordered", "Current Commit Date", "Resale Price",
                       "MPQ", "Invoice Number", "Weight(grams)", "Date Code",
                       "Ship mode", "PGr", "Buyer Name"]
    filtered_df = filtered_df[desired_columns]

    output_file_path = get_output_path(file_name)
    filtered_df.to_excel(output_file_path, index=False)

    buyer_name_col = filtered_df.filter(like='Buyer Name', axis=1).columns[0]
    unique_buyers = filtered_df[buyer_name_col].unique()
    for name in unique_buyers:
        print(name)

if __name__ == "__main__":
    try:
        main()
        input("Complete press Enter to exit")
    except Exception:
        import traceback
        traceback.print_exc()
        input("Program crashed; press Enter to exit")
