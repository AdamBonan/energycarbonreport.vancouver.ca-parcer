import requests
import pandas as pd
import argparse
import time


def write_to_excel(file_name, data, sheet_name="Sheet1"):
    headers = ["Report year", "Building ID", "Use type", "Reported GFA (square feet)", "Address", "Postal code", "Compliance status"]
    df_headers = pd.DataFrame(columns=headers)

    try:
        df_headers.to_excel(file_name, sheet_name=sheet_name, index=False)
        print(True)
    except FileNotFoundError:
        pass

    df_data = pd.DataFrame(data, columns=headers)

    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df_data.to_excel(writer, sheet_name=sheet_name, index=False, header=False,
                         startrow=writer.sheets[sheet_name].max_row)


def get_ids() -> list[list]:

    building_ids = []
    building_ids_chunk10 = []
    for i in range(1, 1501):
        building_ids.append(f"V{10000+i}")

        if len(building_ids) == 10:
            building_ids_chunk10.append(building_ids)
            building_ids = []

    if building_ids:
        building_ids_chunk10.append(building_ids)

    return building_ids_chunk10


def main(file_name: str, time_sleep: int):

    chunk_data = []
    n = 0
    for chunk in get_ids():
        building_ids = ",".join(chunk)

        url = "https://app.touchstoneiq.com/api/buildingid/lookup-id/vancouver"
        data = {
            "building_id": building_ids
        }
        headers = {
            'Accept': 'application/json, text/plain, */*',
            'Accept-Encoding': 'gzip, deflate, br, zstd',
            'Content-Type': 'application/json',
            'Origin': 'https://energycarbonreport.vancouver.ca',
            'Referer': 'https://energycarbonreport.vancouver.ca/',
            'Sec-Ch-Ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'
        }

        array_data = requests.post(url, headers=headers, json=data).json()
        sorted_data = sorted(array_data, key=lambda x: x['custom_building_id'])

        for array in sorted_data:
            street = array.get("street", "").title()
            city = array.get("city", "").title()
            state = array.get("state", "").upper()

            suite = array.get("suite", None)
            if not suite:
                suite = "N/A"

            chunk_data.append([
                array.get("reporting_start_date", None),            # Report year
                array.get("custom_building_id", None),              # Building ID
                array.get("usetype", None).capitalize(),            # Use type
                suite,                                              # Reported GFA (square feet)
                f"{street}, {city}, {state}",                       # Address
                array.get("zipcode", None),                         # Postal code
                array.get("status", None).capitalize()              # Compliance status
            ])

        time.sleep(time_sleep)

    write_to_excel(f"{file_name}.xlsx", chunk_data)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("name", type=str, help="Name exel file")
    parser.add_argument("--sleep", type=int, help="Sleep if blocked requests, (3 sec)", default=0)

    args = parser.parse_args()

    main(args.name, args.sleep)