# Install required libraries
!pip install pandas matplotlib seaborn openpyxl

# Import necessary libraries
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from google.colab import files

# Upload your Excel file
print("Please upload your Excel file:")
uploaded = files.upload()

# Get the filename of the uploaded file
file_name = list(uploaded.keys())[0]

# Specify the sheet name
sheet_to_use = "Sheet1"

def read_and_plot_incidents_by_store(file_name, sheet_name, selected_stores=None):
    try:
        # Read the Excel file with specific sheet
        df = pd.read_excel(file_name, sheet_name=sheet_name)

        # Display basic information about the dataset
        print(f"Dataset Info from sheet '{sheet_name}':")
        print(df.info())
        print("\nFirst few rows:")
        print(df.head())

        # Ensure Column B (dates) is in datetime format
        df['Date'] = pd.to_datetime(df.iloc[:, 1])  # Column B is index 1

        # Filter data from July 2024 onwards
        start_date = pd.to_datetime("2024-07-01")
        df = df[df['Date'] >= start_date]

        # Extract month name from dates
        df['Month'] = df['Date'].dt.strftime('%B')  # e.g., "July", "August"

        # Group by Month and Store (Column G, index 6), count incidents
        incidents_by_store = df.groupby(['Month', df.iloc[:, 6]]).size().unstack(fill_value=0)

        # Define months in order for plotting and summary
        month_order = ['July', 'August', 'September', 'October', 'November', 'December','January', 'February', 'March']
        incidents_by_store = incidents_by_store.reindex(month_order, fill_value=0)

        # Filter to only selected stores if provided
        if selected_stores is not None:
            incidents_by_store = incidents_by_store[selected_stores]

        # Display summary table
        print("\nSummary of Incidents by Month and Store:")
        print(incidents_by_store)

        # Save summary to Excel
        output_file = 'incidents_summary.xlsx'
        incidents_by_store.to_excel(output_file, index=True)
        print(f"\nSummary saved to {output_file}")
        files.download(output_file)

        # Find peak values and months for each store
        peak_info = {}
        for store in incidents_by_store.columns:
            max_incidents = incidents_by_store[store].max()
            peak_month = incidents_by_store[store].idxmax()
            peak_info[store] = (peak_month, max_incidents)

        # Display peak information
        print("\nPeak Values for Each Store:")
        for store, (month, value) in peak_info.items():
            print(f"{store}: {value} incidents in {month}")

        # Define colors for each store
        store_colors = {
          'MTR': 'blue',    # Replace with your store names and colors
            'TAU': 'red',
            'HAM': 'green',
            'HAW': 'yellow',
           'WLG': 'purple'
        }

        # Create a figure with transparent background
        fig = plt.figure(figsize=(12, 6))
        fig.set_facecolor('none')

        # Add axes with transparent background
        ax = plt.gca()
        ax.set_facecolor('none')

        # Plot a line for each selected store and annotate peaks
        for store in incidents_by_store.columns:
            color = store_colors.get(store, 'gray')
            plt.plot(incidents_by_store.index, incidents_by_store[store],
                    marker='o', linestyle='-', label=store, color=color)
            # Annotate peak
            peak_month, peak_value = peak_info[store]
            plt.annotate(f'{peak_value}',
                        xy=(peak_month, peak_value),
                        xytext=(0, 5),  # Offset above the point
                        textcoords='offset points',
                        ha='center',
                        color=color)

        # Customize the plot
        plt.xlabel('Month')
        plt.ylabel('Number of Incidents')
        plt.title('Number of Incidents Per Month by Store (July 2024 Onwards)')
        plt.xticks(rotation=45)
        plt.grid(True, linestyle='--', alpha=0.7)

        # Customize axis line colors
        ax.spines['bottom'].set_color('black')
        ax.spines['left'].set_color('black')
        ax.spines['top'].set_color('none')
        ax.spines['right'].set_color('none')

        # Customize the legend
        plt.legend(
            title='Store',
            bbox_to_anchor=(1.05, 1),
            loc='upper left',
            fontsize=10,
            title_fontsize=12,
            facecolor='white',
            edgecolor='black'
        )

        # Adjust layout and display
        plt.tight_layout()
        plt.show()

    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Specify the stores you want to visualize
selected_stores = ['MTR', 'TAU', 'HAM','HAW','WLG']  # Replace with your desired store names

# Run the visualization with the specified sheet and selected stores
read_and_plot_incidents_by_store(file_name, sheet_name=sheet_to_use, selected_stores=selected_stores)
