#!/usr/bin/env python3

import pandas
import argparse


parser = argparse.ArgumentParser()
parser.add_argument("filename", help="The spreadsheet to open, in xls format")
parser.add_argument(
    "--sheet-name",
    help="The name of the sheet containing the information to extract",
    default="Device Configurations",
)
parser.add_argument(
    "--ip-row-name",
    help="The name of the ip address row within the sheet in order to populate the second column",
    default="OOB MGMT IP",
)
parser.add_argument(
    "--hostname-row-name",
    help="The name of the hostname row with the sheet in order to populate the first column",
    default="Hostname",
)
parser.add_argument(
    "--device-type-row-name",
    help="The name of the device ID row",
    default="Device Model",
)
parser.add_argument(
    "--device-type", help="The device ID to filter for", default=[], action="append"
)
parser.add_argument(
    "--header-row",
    help="The row ID of the column names row, 0 indexed",
    default=3,
    type=int,
)

args = parser.parse_args()

with open(args.filename, "rb") as file:
    data = pandas.read_excel(
        file,
        sheet_name=args.sheet_name,
        header=args.header_row,
        usecols=[args.hostname_row_name, args.ip_row_name, args.device_type_row_name],
    )
    for i in data.index:
        if data[args.device_type_row_name][i] in args.device_type and (
            not pandas.isna(data[args.hostname_row_name][i])
            and not pandas.isna(data[args.ip_row_name][i])
        ):
            print(f"{data[args.hostname_row_name][i]},{data[args.ip_row_name][i]}")
