import sys
import json
import pandas as pd

def process_dataframe(df, sheet_name):
    # Check if the required columns exist
    if "Tag" not in df.columns or "Example Data" not in df.columns:
        print(f"Error: Required columns not found in the input data for sheet '{sheet_name}'. Skipping this sheet.")
        return
    data = {
        "Device": {
            "name": df.loc[df["Tag"] == "Device Name"]["Example Data"].values[0],
            "model number": df.loc[df["Tag"] == "Device Model Number"]["Example Data"].values[0],
            "category": df.loc[df["Tag"] == "Device Category"]["Example Data"].values[0],
            "identifier": df.loc[df["Tag"] == "Device Identifier"]["Example Data"].values[0],
            "cooling": df.loc[df["Tag"] == "Device Part SKU"]["Example Data"].values[0],
            "power": df.loc[df["Tag"] == "Device Part SKU"]["Example Data"].values[0],
            "powerRedundancy": df.loc[df["Tag"] == "Device Cooling"]["Example Data"].values[0],
            "shelfCountMax": df.loc[df["Tag"] == "Device Shelf Count -max"]["Example Data"].values[0],
            "shelfCountFound": df.loc[df["Tag"] == "Device Shelf Count - found"]["Example Data"].values[0],
            "Shelves": []
        }
    }

    shelves = df.loc[df["Tag"] == "Device Shelf Count -max"]["Example Data"].values[0]
    try:
        shelves = int(shelves)
    except ValueError:
        shelves = 0

    for i in range(1, shelves + 1):
        shelf = {
            "count": i,
            "ioCardCountMax": df.loc[df["Tag"] == "Shelf I/O Card Count - max"]["Example Data"].values[0],
            "ioCardCountFound": df.loc[df["Tag"] == "Shelf I/O Card Count - found"]["Example Data"].values[0],
            "Cards": []
        }

        cards = df.loc[df["Tag"] == "Shelf I/O Card Count - max"]["Example Data"].values[0]
        try:
            cards = int(cards)
        except ValueError:
            cards = 0

        for j in range(1, cards + 1):
            card = {
                "Card Number": j,
                "Card Type": df.loc[df["Tag"] == "Card Type"]["Example Data"].values[0],
                "Card Model": df.loc[df["Tag"] == "Card Model"]["Example Data"].values[0],
                "Card SKU": df.loc[df["Tag"] == "Card SKU"]["Example Data"].values[0],
                "Card Sub-Card Count - found": df.loc[df["Tag"] == "Card Sub-Card Count - found"]["Example Data"].values[0],
                "Sub-Cards": []
            }

            sub_cards = df.loc[df["Tag"] == "Card Sub-Card Count - found"]["Example Data"].values[0]
            try:
                sub_cards = int(sub_cards)
            except ValueError:
                sub_cards = 0

            for k in range(1, sub_cards + 1):
                sub_card = {
                    "Sub-Card Number": k,
                    "Sub-Card Type": df.loc[df["Tag"] == "Sub-Card Type"]["Example Data"].values[0],
                    "Sub-Card Model": df.loc[df["Tag"] == "Sub-Card Model"]["Example Data"].values[0],
                    "Sub-Card SKU": df.loc[df["Tag"] == "Sub-Card SKU"]["Example Data"].values[0],
                    "Sub-Card Port Count - Found": df.loc[df["Tag"] == "Sub-Card Port Count - Found"]["Example Data"].values[0],
                    "Ports": []
                }

                ports = df.loc[df["Tag"] == "Sub-Card Port Count - Found"]["Example Data"].values[0]
                try:
                    ports = int(ports)
                except ValueError:
                    ports = 0

                for l in range(1, ports + 1):
                    port = {
                        "Sub-Card Port Number": l,
                        "Sub-Card Port Number Speed": df.loc[df["Tag"] == "Sub-Card Port Number Speed"]["Example Data"].values[0],
                        "Sub-Card Port Connection": df.loc[df["Tag"] == "Sub-Card Port Connection"]["Example Data"].values[0],"Sub-Card Port Number Status": df.loc[df["Tag"] == "Sub-Card Port Number Status"]["Example Data"].values[0]
                    }

                    sub_card["Ports"].append(port)
                card["Sub-Cards"].append(sub_card)
            shelf["Cards"].append(card)
        data["Device"]["Shelves"].append(shelf)

    return data

def xlsx_to_json(file_path):
    try:
        xls = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        for sheet_name, df in xls.items():
            json_file_name = f"{sheet_name}.json"
            processed_data = process_dataframe(df, sheet_name)
            with open(json_file_name, 'w') as json_file:
                json.dump(processed_data, json_file, indent=2)
            print(f"Created JSON file: {json_file_name}")
    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python xlsx_to_json.py <file_path>")
        sys.exit(1)

    file_path = sys.argv[1]
    xlsx_to_json(file_path)