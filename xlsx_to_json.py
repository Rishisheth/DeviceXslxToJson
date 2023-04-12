import sys
import json
import pandas as pd

def process_dataframe(df, sheet_name):
    # Check if the required columns exist
    if "Tag" not in df.columns or "Example Data" not in df.columns:
        print(f"Error: Required columns not found in the input data for sheet '{sheet_name}'. Skipping this sheet.")
        return None
    data = {
        "device": {
            "name": df.loc[df["Tag"] == "Device Name"]["Example Data"].values[0],
            "model number": df.loc[df["Tag"] == "Device Model Number"]["Example Data"].values[0],
            "category": df.loc[df["Tag"] == "Device Category"]["Example Data"].values[0],
            "identifier": df.loc[df["Tag"] == "Device Identifier"]["Example Data"].values[0],
            "partSku": df.loc[df["Tag"] == "Device Part SKU"]["Example Data"].values[0],
            "comments": {
                "comment": []
            },
            "cooling": df.loc[df["Tag"] == "Device Cooling"]["Example Data"].values[0],
            "power": df.loc[df["Tag"] == "Device Part SKU"]["Example Data"].values[0],
            "powerRedundancy": df.loc[df["Tag"] == "Device Cooling"]["Example Data"].values[0],
            "shelfCountMax": df.loc[df["Tag"] == "Device Shelf Count -max"]["Example Data"].values[0],
            "shelfCountFound": df.loc[df["Tag"] == "Device Shelf Count - found"]["Example Data"].values[0],
            "shelves": {
                "shelf": []
            }
        }
    }

    comments = {
        "dateCreated": df.loc[df["Tag"] == "Device Comment - Date"]["Example Data"].values[0],
        "userId": df.loc[df["Tag"] == "Device Comment - UserID"]["Example Data"].values[0],
        "text": df.loc[df["Tag"] == "Device Comment - Text"]["Example Data"].values[0],
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
            "ioSubCardCountMax": df.loc[df["Tag"] == "Shelf I/O Card - SubCard Count - max"]["Example Data"].values[0],
            "ioSubCardCountFound": df.loc[df["Tag"] == "Shelf I/O Card - SubCard Count -found"]["Example Data"].values[0],
            "otherCardCountFound": df.loc[df["Tag"] == "Shelf Other Card Count - found"]["Example Data"].values[0],
            "otherCardsubCardCountFound": df.loc[df["Tag"] == "Shelf Other Card - SubCard Count -found"]["Example Data"].values[0],
            "metroENetwork": df.loc[df["Tag"] == "Shelf MEN ID"]["Example Data"].values[0],
            "metroENetwork": df.loc[df["Tag"] == "Shelf CO ID"]["Example Data"].values[0],
            "ringId": df.loc[df["Tag"] == "Shelf Ring ID"]["Example Data"].values[0],
            "lATA": df.loc[df["Tag"] == "Shelf LATA"]["Example Data"].values[0],
            "status": df.loc[df["Tag"] == "Shelf Status"]["Example Data"].values[0],
            "comments": {
                "comment": [comments]
            },
            "cards": {
                "card": []
            }
        }

        cards = df.loc[df["Tag"] == "Shelf I/O Card Count - max"]["Example Data"].values[0]
        try:
            cards = int(cards)
        except ValueError:
            cards = 0

        for j in range(1, cards + 1):
            card = {
                "number": j,
                "type": df.loc[df["Tag"] == "Card Type"]["Example Data"].values[0],
                "model": df.loc[df["Tag"] == "Card Model"]["Example Data"].values[0],
                "sku": df.loc[df["Tag"] == "Card SKU"]["Example Data"].values[0],
                "status": df.loc[df["Tag"] == "Card Status"]["Example Data"].values[0],
                "comments": {
                    "comment": [comments]
                },
                "subCardCountFound": df.loc[df["Tag"] == "Card Sub-Card Count - found"]["Example Data"].values[0],
                "subCards": {
                    "subCard": []
                },
                "ports": {
                    "port": []
                },
                "services": {
                    "service": []
                }
            }

            sub_cards = df.loc[df["Tag"] == "Card Sub-Card Count - found"]["Example Data"].values[0]
            try:
                sub_cards = int(sub_cards)
            except ValueError:
                sub_cards = 0

            ports = df.loc[df["Tag"] == "Sub-Card Port Count - Found"]["Example Data"].values[0]
            try:
                ports = int(ports)
            except ValueError:
                ports = 0

            services = {
                    "circuitID": df.loc[df["Tag"] == "CKID"]["Example Data"].values[0],
                    "type": df.loc[df["Tag"] == "CKID Type"]["Example Data"].values[0],
                    "relatedCircuitID": df.loc[df["Tag"] == "Sub-Card Port Number Status"]["Example Data"].values[0],
                    "identifier": df.loc[df["Tag"] == "Device Identifier.1"]["Example Data"].values[0],
                    "evcid": df.loc[df["Tag"] == "EVC ID.2"]["Example Data"].values[0],
                }
            
            for l in range(1, ports + 1):
                port = {
                    "number": l,
                    "connection": df.loc[df["Tag"] == "Sub-Card Port Connection"]["Example Data"].values[0],
                    "status": df.loc[df["Tag"] == "Sub-Card Port Number Status"]["Example Data"].values[0],
                }

            for k in range(1, sub_cards + 1):
                sub_card = {
                    "number": k,
                    "type": df.loc[df["Tag"] == "Sub-Card Type"]["Example Data"].values[0],
                    "model": df.loc[df["Tag"] == "Sub-Card Model"]["Example Data"].values[0],
                    "sku": df.loc[df["Tag"] == "Sub-Card SKU"]["Example Data"].values[0],
                    "status": df.loc[df["Tag"] == "Sub-Card Status"]["Example Data"].values[0],
                    "comments": {
                        "comment": [comments]
                    },
                    "portCount": df.loc[df["Tag"] == "Sub-Card Port Count - Found"]["Example Data"].values[0]
                }

                card["services"]["service"].append(services)
                card["ports"]["port"].append(port)
                card["subCards"]["subCard"].append(sub_card)
            shelf["cards"]["card"].append(card)
        data["device"]["comments"]["comment"].append(comments)
        data["device"]["shelves"]["shelf"].append(shelf)

    return data

def xlsx_to_json(file_path):
    try:
        xls = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        for sheet_name, df in xls.items():
            json_file_name = f"{sheet_name}.json"
            if sheet_name.__contains__("ObjMOd"):
                processed_data = process_dataframe(df, sheet_name)
                if processed_data is not None:
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