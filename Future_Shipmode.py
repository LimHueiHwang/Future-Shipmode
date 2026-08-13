import datetime
import os

import pandas as pd


# ============================================================
# Configuration
# ============================================================

BASE_PATH = (
    r"//sgsind0nsifsv01a/IMAC Data/"
    r"IMAC Senior or Teams/Europe & Other Asia Team/VN01/"
)

FUTURE_SHIPMODE_FOLDER = os.path.join(
    BASE_PATH,
    "FUTURE SHIP MODE/",
)

VN01_ON_HAND_FOLDER = os.path.join(
    BASE_PATH,
    "VN01 on hand part/",
)


# Required columns from each source file
REQUIRED_OH_COLUMNS = {
    "Material",
    "PGr",
}

REQUIRED_FUTURE_COLUMNS = {
    "Customer Part.",
}


# Final output columns
OUTPUT_COLUMNS = [
    "Ship To",
    "Order Number",
    "Line Number",
    "Customer Part.",
    "Mfr",
    "Mfr Part.",
    "Qty Required",
    "Qty Allocated",
    "Date Ordered",
    "Current Commit Date",
    "Resale Price",
    "MPQ",
    "Invoice Number",
    "Weight(grams)",
    "Date Code",
    "Ship mode",
    "PGr",
    "Buyer Name",
]


# ============================================================
# Date / Path Helpers
# ============================================================

def get_current_iso_week():
    """
    Return the current ISO week in the format WW'YY.

    Example:
        WW33'26
    """
    current_date = datetime.datetime.now()

    iso_year, iso_week, _ = current_date.isocalendar()

    return f"WW{iso_week:02d}'{str(iso_year)[-2:]}"


def get_output_path(file_name):
    """
    Return the output path for the consolidated report.
    """
    current_week = get_current_iso_week()

    return os.path.join(
        FUTURE_SHIPMODE_FOLDER,
        current_week,
        f"JB {file_name}_NEW.xlsx",
    )


def get_oh_part_path():
    """
    Return the VN01 On-Hand report path for the current week.

    The VN01 On-Hand file is named using the Monday
    date of the current week in MMDD format.
    """
    today = datetime.datetime.now()

    monday = today - datetime.timedelta(
        days=today.weekday()
    )

    monday_str = monday.strftime("%m%d")

    return os.path.join(
        VN01_ON_HAND_FOLDER,
        f"VN01 OH {monday_str}.xlsx",
    )


def get_future_shipmode_path(file_name):
    """
    Return the Future Ship Mode input file path.
    """
    current_week = get_current_iso_week()

    return os.path.join(
        FUTURE_SHIPMODE_FOLDER,
        current_week,
        f"{file_name}.xlsx",
    )


# ============================================================
# Validation Helpers
# ============================================================

def validate_file_exists(file_path, description):
    """
    Verify that an input file exists.
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(
            f"{description} was not found:\n"
            f"{file_path}"
        )


def validate_columns(
    dataframe,
    required_columns,
    description,
):
    """
    Verify that required columns exist
    in a dataframe.
    """
    missing_columns = sorted(
        required_columns - set(dataframe.columns)
    )

    if missing_columns:
        raise ValueError(
            f"{description} is missing required "
            f"column(s):\n"
            + "\n".join(
                f"- {column}"
                for column in missing_columns
            )
        )


# ============================================================
# Data Loading
# ============================================================

def load_vn01_on_hand(file_path):
    """
    Load and prepare the VN01 On-Hand report.

    VN01 contains:
        Material
        PGr

    Material is renamed to:
        Customer Part.
    """
    df = pd.read_excel(file_path)

    # Clean Excel column headers
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    validate_columns(
        df,
        REQUIRED_OH_COLUMNS,
        "VN01 On-Hand report",
    )

    # Material is the matching key used
    # against Future Ship Mode Customer Part.
    df = df.rename(
        columns={
            "Material": "Customer Part."
        }
    )

    df["Customer Part."] = (
        df["Customer Part."]
        .astype(str)
        .str.strip()
    )

    # PGr comes from VN01 On-Hand.
    df["PGr"] = (
        df["PGr"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    return df


def load_future_shipmode(file_path):
    """
    Load and prepare the Future Ship Mode report.

    Future Ship Mode contains:
        Customer Part.

    Future Ship Mode does NOT provide PGr.
    """
    df = pd.read_excel(file_path)

    # Clean Excel column headers
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    validate_columns(
        df,
        REQUIRED_FUTURE_COLUMNS,
        "Future Ship Mode report",
    )

    df["Customer Part."] = (
        df["Customer Part."]
        .astype(str)
        .str.strip()
    )

    return df


# ============================================================
# Data Processing
# ============================================================

def merge_and_filter(df_oh, df_sm):
    """
    Merge VN01 On-Hand and Future Ship Mode
    using Customer Part.

    PGr comes from the VN01 On-Hand report.

    After the merge, keep only records where
    PGr starts with 'U'.
    """

    # --------------------------------------------------------
    # Merge
    # --------------------------------------------------------
    merged_df = pd.merge(
        df_oh,
        df_sm,
        on="Customer Part.",
        how="inner",
    )

    if merged_df.empty:
        raise ValueError(
            "No matching parts were found between "
            "the VN01 On-Hand and Future Ship Mode "
            "reports."
        )

    # --------------------------------------------------------
    # Filter
    # --------------------------------------------------------
    filtered_df = merged_df[
        merged_df["PGr"].str.startswith("U")
    ].copy()

    if filtered_df.empty:
        raise ValueError(
            "Matching parts were found, but none "
            "belong to a Purchasing Group starting "
            "with 'U'."
        )

    # --------------------------------------------------------
    # Validate output columns
    # --------------------------------------------------------
    missing_output_columns = [
        column
        for column in OUTPUT_COLUMNS
        if column not in filtered_df.columns
    ]

    if missing_output_columns:
        raise ValueError(
            "The merged data is missing required "
            "output column(s):\n"
            + "\n".join(
                f"- {column}"
                for column in missing_output_columns
            )
        )

    # --------------------------------------------------------
    # Keep only required output columns
    # --------------------------------------------------------
    return filtered_df[
        OUTPUT_COLUMNS
    ].copy()


# ============================================================
# Output
# ============================================================

def save_output(dataframe, output_path):
    """
    Save the consolidated report to the server.
    """
    output_directory = os.path.dirname(
        output_path
    )

    if not os.path.isdir(output_directory):
        raise FileNotFoundError(
            "Output folder was not found:\n"
            f"{output_directory}"
        )

    dataframe.to_excel(
        output_path,
        index=False,
    )


def print_buyer_list(dataframe):
    """
    Print unique Buyer Names for reference.
    """
    buyer_names = (
        dataframe["Buyer Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    buyer_names = sorted(
        name
        for name in buyer_names.unique()
        if name
    )

    print("\nBuyer Names:")
    print("-" * 30)

    if not buyer_names:
        print("No buyer names found.")
        return

    for buyer in buyer_names:
        print(buyer)


# ============================================================
# Main Workflow
# ============================================================

def main():

    # --------------------------------------------------------
    # Get Future Ship Mode file name
    # --------------------------------------------------------
    file_name = input(
        "Enter Future Ship Mode file name: "
    ).strip()

    if not file_name:
        raise ValueError(
            "A Future Ship Mode file name is required."
        )

    # --------------------------------------------------------
    # Build file paths
    # --------------------------------------------------------
    future_shipmode_path = (
        get_future_shipmode_path(
            file_name
        )
    )

    vn01_on_hand_path = (
        get_oh_part_path()
    )

    print("\nProcessing files:")
    print(
        f"VN01 On-Hand     : "
        f"{vn01_on_hand_path}"
    )
    print(
        f"Future Ship Mode : "
        f"{future_shipmode_path}"
    )

    # --------------------------------------------------------
    # Validate input files
    # --------------------------------------------------------
    validate_file_exists(
        vn01_on_hand_path,
        "VN01 On-Hand report",
    )

    validate_file_exists(
        future_shipmode_path,
        "Future Ship Mode report",
    )

    # --------------------------------------------------------
    # Load input reports
    # --------------------------------------------------------
    df_oh = load_vn01_on_hand(
        vn01_on_hand_path
    )

    df_sm = load_future_shipmode(
        future_shipmode_path
    )

    # --------------------------------------------------------
    # Merge and apply business rule
    # --------------------------------------------------------
    filtered_df = merge_and_filter(
        df_oh,
        df_sm,
    )

    # --------------------------------------------------------
    # Save consolidated report
    # --------------------------------------------------------
    output_file_path = (
        get_output_path(
            file_name
        )
    )

    save_output(
        filtered_df,
        output_file_path,
    )

    # --------------------------------------------------------
    # Print Buyer list
    # --------------------------------------------------------
    print_buyer_list(
        filtered_df
    )

    # --------------------------------------------------------
    # Completion message
    # --------------------------------------------------------
    print(
        "\nProcessing completed successfully."
    )

    print(
        f"Output file: "
        f"{output_file_path}"
    )


# ============================================================
# Application Entry Point
# ============================================================

if __name__ == "__main__":

    try:

        main()

        input(
            "\nComplete. Press Enter to exit."
        )

    except Exception as error:

        print("\nAutomation failed.")
        print("-" * 30)
        print(error)

        input(
            "\nPress Enter to exit."
        )
