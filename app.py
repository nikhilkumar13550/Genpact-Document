import os
import json
from datetime import datetime
import pandas as pd

# -----------------------------
# CONFIG
# -----------------------------
INPUT_FOLDER = "json_input"
OUTPUT_FILE = "output.xlsx"


# -----------------------------
# Process Single JSON File
# -----------------------------
def process_json_file(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    formatted_row = {}

    for key, inner_dict in data.items():
        value = inner_dict.get("value")
        confidence = inner_dict.get("confidence")

        formatted_row[key] = f"{value},{confidence}"

    formatted_row["Processed_Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    return formatted_row


# -----------------------------
# Move File Safely
# -----------------------------
def move_file(file_path, processed_folder):
    os.makedirs(processed_folder, exist_ok=True)

    filename = os.path.basename(file_path)
    new_path = os.path.join(processed_folder, filename)

    # If file already exists in processed folder
    if os.path.exists(new_path):
        base, ext = os.path.splitext(filename)
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        new_path = os.path.join(processed_folder, f"{base}_{timestamp}{ext}")

    os.rename(file_path, new_path)


# -----------------------------
# Main Logic
# -----------------------------
def main():
    if not os.path.exists(INPUT_FOLDER):
        print("json_input folder not found.")
        return

    processed_folder = os.path.join(INPUT_FOLDER, "processed")
    all_rows = []

    for file in os.listdir(INPUT_FOLDER):
        if file.lower().endswith(".json"):
            full_path = os.path.join(INPUT_FOLDER, file)
            print(f"Processing: {file}")

            try:
                row = process_json_file(full_path)
                all_rows.append(row)
                move_file(full_path, processed_folder)

            except Exception as e:
                print(f"Error processing {file}: {e}")

    if not all_rows:
        print("No JSON files found.")
        return

    new_df = pd.DataFrame(all_rows)

    if os.path.exists(OUTPUT_FILE):
        try:
            existing_df = pd.read_excel(OUTPUT_FILE)
            final_df = pd.concat([existing_df, new_df], ignore_index=True)
        except:
            final_df = new_df
    else:
        final_df = new_df

    final_df.to_excel(OUTPUT_FILE, index=False)

    print("Excel updated successfully!")


if __name__ == "__main__":
    main()