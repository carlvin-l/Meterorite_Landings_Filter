from meteor_data_class import MeteorDataEntry
from datetime import datetime
import pandas as pd

# -----------------------------
# Filter function
# -----------------------------
def filter_data(data, attribute, lower_limit, upper_limit):
    filtered_data = []
    for entry in data:
        if attribute == "mass" and entry.mass.isdigit():
            if lower_limit <= float(entry.mass) <= upper_limit:
                filtered_data.append(entry)
        elif attribute == "year" and entry.year.isdigit():
            if lower_limit <= int(entry.year) <= upper_limit:
                filtered_data.append(entry)
    return filtered_data

# -----------------------------
# Excel export function
# -----------------------------
def save_to_excel(filtered_data, filename=None):
    # Ensure filtered_data is not empty
    if not filtered_data:
        print("No data to save.")
        return

    # Prepare list of dictionaries
    data_dicts = []
    for entry in filtered_data:
        data_dicts.append({
            "NAME": entry.name,
            "ID": entry.id,
            "NAMETYPE": entry.nameType,
            "RECCLASS": entry.reClass,
            "MASS (g)": entry.mass,
            "FALL": entry.fall,
            "YEAR": entry.year,
            "RECLAT": entry.reClat,
            "RECLONG": entry.reClong,
            "GEOLOCATION": entry.geoLocation,
            "STATES": entry.states,
            "COUNTIES": entry.counties
        })

    # Create DataFrame
    columns = ["NAME", "ID", "NAMETYPE", "RECCLASS", "MASS (g)", "FALL",
               "YEAR", "RECLAT", "RECLONG", "GEOLOCATION", "STATES", "COUNTIES"]
    df = pd.DataFrame(data_dicts, columns=columns)

    # Generate filename if none provided
    if not filename:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"filtered_meteor_data_{timestamp}.xlsx"

    # Ensure the filename ends with .xlsx
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"

    # Make sure pandas uses openpyxl engine explicitly
    try:
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"\nFiltered data saved to '{filename}'")
    except ImportError:
        print("Error: openpyxl module is required to write Excel files.")
        print("Install it via: pip install openpyxl")
# -----------------------------
# Main program
# -----------------------------
def main():
    print("\nWelcome to the meteorite filtering program!\n"
          "Filters meteorite landings by mass or year.\n")

    # Input file
    user_input1 = input("\nEnter a valid file name (ex. 'file_name.txt') or 'q' to quit: ")
    if user_input1.lower() == 'q':
        print('\nExiting program. Peace!')
        return

    try:
        # Read file and parse entries
        with open(user_input1, "r") as file_obj:
            data = []
            next(file_obj)  # skip header
            for line in file_obj:
                splitted = line.strip('\n').split('\t')
                while len(splitted) < 12:  # 12 = number of fields in MeteorDataEntry
                    splitted.append("")
                data.append(MeteorDataEntry(*splitted))

        # Select filter type
        print("\nFilter by:\n1. Mass (g)\n2. Year\n3. Quit")
        choice = input(">> ")
        if choice == "1":
            lower = input("Enter LOWER mass (g) ('Q' to quit): ")
            if lower.upper() == 'Q': return
            upper = input("Enter UPPER mass (g) ('Q' to quit): ")
            if upper.upper() == 'Q': return
            filtered_data = filter_data(data, "mass", float(lower), float(upper))
        elif choice == "2":
            lower = input("Enter LOWER year ('Q' to quit): ")
            if lower.upper() == 'Q': return
            upper = input("Enter UPPER year ('Q' to quit): ")
            if upper.upper() == 'Q': return
            filtered_data = filter_data(data, "year", int(lower), int(upper))
        else:
            print("Exiting program. Peace!")
            return

        # Print a summary table
        print(f"\nFiltered {len(filtered_data)} entries.")
        for i, entry in enumerate(filtered_data, start=1):
            print(f"{i:<4}{entry.name:<15}{entry.mass:<10}{entry.year:<6}")

        # Prompt to save to Excel
        save_choice = input("\nSave filtered data to Excel? (Y/N): ")
        if save_choice.lower() == 'y':
            user_filename = input("Enter filename (leave blank for auto-generated): ")
            save_to_excel(filtered_data, filename=user_filename if user_filename else None)
        else:
            print("Excel export skipped.")

    except FileNotFoundError:
        print("\nFile not found. Exiting...")

if __name__ == "__main__":
    main()