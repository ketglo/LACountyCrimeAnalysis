import os
import pandas as pd

# Set display options to avoid truncation
pd.set_option("display.max_colwidth", None)  # Display full column content
pd.set_option("display.max_rows", None)  # Display all rows

script_dir = os.path.dirname(os.path.realpath(__file__))

def list_files_in_current_folder():
    try:
        current_directory = os.getcwd()
        files = [f for f in os.listdir(current_directory) if os.path.isfile(f)]
        return files
    except Exception as e:
        print(f"An error occurred: {e}")
        return []

def read_all_sheets_from_excel(filename) -> dict:
    try:
        print(filename)
        path = os.path.join(script_dir, filename)
        print(path)
        xls = pd.ExcelFile(path)
        df_dict = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}
        return df_dict
    except FileNotFoundError:
        print(f"File '{filename}' not found. Make sure it's in the same folder as the Python script.")
        return {}

# Example usage
if __name__ == "__main__":
    file_list = list_files_in_current_folder()
    dataframes = ''
    if not file_list:
        print("No files found in the current folder.")
    else:
        print("Files in the current folder:")
        for i, file in enumerate(file_list, start=1):
            print(f"{i}. {file}")
        selected_file_index = input("Enter the number corresponding to the Excel file you want to read: ")
        try:
            selected_file_index = int(selected_file_index)
            if 1 <= selected_file_index <= len(file_list):
                selected_filename = file_list[selected_file_index - 1]
                dataframes = read_all_sheets_from_excel(selected_filename)
                if not dataframes:
                    print(f"No data loaded from '{selected_filename}'.")
                else:
                    print(f"Data loaded successfully from '{selected_filename}'!")
            else:
                print("Invalid selection. Please choose a valid file number.")
        except ValueError:
            print("Invalid input. Please enter a valid file number.")
    if not dataframes:
        print("No data loaded due to missing Excel file.")
    else:
        print("Data loaded successfully!")
        print("What information do you want to Analyze?")
        print("1. Top Crimes per City")
        sortby = input("")

        # print each dataframe name
        #print("Dataframe keys of dataframes:" + ", ".join(dataframes.keys()))

        for k, v in dataframes.items():
            # strip whitespace where possible from column names; need to check if isinstance(x, str) because some column names are numbers
            try:
                v = v.rename(columns=lambda x: x.strip() if isinstance(x, str) else x)
            except:
                pass

            # strip whitespace where possible from cells
            try:
                v = v.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
            except:
                pass
            dataframes[k] = v
            #print('dataframe: '+ k)
            #print(v.head())


        # Start with aggregating the data by city
        df = dataframes['PART_I_AND_II_CRIMES_YTD_0']

        # Count total crimes by city
        total_crimes_by_city = df['City'].value_counts().rename_axis('City').reset_index(name='Total Crimes')

        # Count gang-related crimes by city
        gang_crimes_by_city = df[df['Gang Related'] == 'YES']['City'].value_counts().rename_axis('City').reset_index(name='Gang Related Crimes')

        # Print the results
        #for index, row in gang_crimes_by_city.iterrows():
            #print(f"City: {row['City']}, Gang Related Crimes: {row['Gang Related Crimes']}")

        # Print unique values in the 'Gang Related' column
        #print("Unique values in 'Gang Related' column:", df['Gang Related'].unique())

        # Find the top crime for each city and its count
        top_crime_by_city = df.groupby('City')['Stat Code Desc'].agg(lambda x: x.value_counts().index[0]).reset_index(name='Top Crime')
        top_crime_count_by_city = df.groupby(['City', 'Stat Code Desc']).size().groupby(level=0).max().reset_index(name='Top Crime Count')

        # Merge the dataframes
        city_crime_stats = total_crimes_by_city.merge(top_crime_by_city, on='City',how='left').merge(top_crime_count_by_city,on='City', how='left')

        # Sort by total crimes and get top 10 cities
        top_10_city_crime_stats = city_crime_stats.sort_values(by='Total Crimes', ascending=False).head(10)

        # Add the 'gang_crimes_by_city' column to the top_10_city_crime_stats dataframe
        top_10_city_crime_stats['Gang Related Crimes'] = gang_crimes_by_city['Gang Related Crimes']
        pd.set_option('display.max_columns', None)

        sort = top_10_city_crime_stats.sort_values(by='Top Crime Count', ascending=False).head(10)

        #print(top_10_city_crime_stats)

        # Save the top_10_city_crime_stats dataframe to an Excel file
        sort.to_excel(os.path.join(script_dir,"top_10_city_crime_stats.xlsx"), index=False)

        print("Data saved to top_10_city_crime_stats.xlsx")
