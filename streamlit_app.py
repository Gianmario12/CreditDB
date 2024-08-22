import pandas as pd
import streamlit as st
import os
import glob
import pythoncom
from st_aggrid import AgGrid
from win32com.client import Dispatch
from io import BytesIO
import tempfile

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from webdriver_manager.chrome import ChromeDriverManager

st.title("🎈 My new app")
################################################################### PACKAGES

# Configure the page
st.set_page_config(page_title="Credit Dashboard",
                   page_icon=":bar_chart:",
                   layout="wide")

# Read the Excel file
#file_path = "C:/Users/GIANMARIO1/OneDrive - OMV Group/Desktop/Python_test/OMV_dashboard.xlsm"
file_path = "//somvat002003/OGP/OGP/SMT/PC/00_DFG-C/03_NEW OGMT CREDIT/17_TESTENV/Dashboard/OMV_dashboard.xlsm"
df = pd.read_excel(file_path, sheet_name="Main", usecols="A:R")

# Filter out the rows where all the values in columns 10 to 17 are 0
df = df.loc[~(df.iloc[:, 10:17] == 0).all(axis=1)]

# Filter out the rows where 'SAP No' column is blank
df = df[df["SAP No"].notnull()]

# Add the breaches
df['Breach'] = df.apply(lambda row: max(row['TOT Credit Exp.'] - row['total_credit_limit'], 0), axis=1)

#Some formatting
columns_to_format = ["95%PFE", "TOT Credit Exp.", "Unsec. Exp.", "MtM", "securities", "unsecured_limit", "total_credit_limit", "SE", "Breach"] 
df[columns_to_format] = df[columns_to_format].applymap(lambda x: f"€{x:,.2f}")

# Remove "IT Errors" from the DataFrame
ITErrors = df[((df["SE"] == df["Unsec. Exp."]) & 
          (df["Unsec. Exp."] == df["TOT Credit Exp."]) & 
          (df["95%PFE"] == "€0.00") & 
          (df["MtM"] == "€0.00") &
          (df["SE"] != "€0.00"))]

df = df[~((df["SE"] == df["Unsec. Exp."]) & 
          (df["Unsec. Exp."] == df["TOT Credit Exp."]) & 
          (df["95%PFE"] == "€0.00") & 
          (df["MtM"] == "€0.00"))]

# Create a new DataFrame with rows where 'Breach' is greater than 0 (for simplicity 1)
breachesdf = df[df['Breach'].apply(lambda x: float(x.replace('€', '').replace(',', '')) > 1)]
breachesdf[columns_to_format] = breachesdf[columns_to_format].applymap(lambda x: float(x.replace('€', '').replace(',', '')))
columns_to_keep = ["customer", "counterparty", "RC", "total_credit_limit", "unsecured_limit", "securities", "Unsec. Exp.", "TOT Credit Exp.", "Breach"]
breachesdf = breachesdf[columns_to_keep]

# Create a new DataFrame for the DUNS update
# Create the DUNSupdate DataFrame with the specified columns
columns = ["Company_Name", "Address_Line_1", "Address_Line_2", "1949", "1976", "State_Province", "Country_Code", "Phone_Number", "Registration_number", 
    "Tax_number", "DUNS_Number", "Folder 2", "Folder 3", "Folder 4", "Folder 5", "Folder 6", "Folder 7", "Folder 8", "Folder 9", "Folder 10", 
    "Order_Reference" ]

# Initialize the DataFrame with empty values
DUNSupdate = pd.DataFrame(columns=columns)

# Populate 'DUNS_Number' and 'Country_Code' from 'db'
DUNSupdate['DUNS_Number'] = df['duns_number']
DUNSupdate['Country_Code'] = df['country']

# Define the mapping dictionary
country_to_folder2 = {
    "GB": "GB companies",
    "FR": "FR companies",
    "BE": "BE companies",
    "AT": "AT companies",
    "CH": "CH companies",
    "DE": "DE companies",
    "IT": "IT companies",
    "HU": "HU companies",
    "SI": "SI companies",
    "CZ": "CZ companies",
    "NL": "NL companies",
    "LU": "LU companies",
    "ES": "ES companies",
    "IE": "IE companies"
}

# Function to map country code to folder 2 value
def map_country_to_folder2(country_code):
    return country_to_folder2.get(country_code, "All CPs in monitoring")

# Apply the function to the 'Country_Code' column to populate 'Folder 2'
DUNSupdate['Folder 2'] = DUNSupdate['Country_Code'].apply(map_country_to_folder2)

# Function to generate the Company_Name based on DUNS_Number
def generate_company_name(duns_number):
    duns_str = str(duns_number)
    return f"{duns_str[:2]}-{duns_str[2:5]}-{duns_str[5:]}"

# Apply the function to the 'DUNS_Number' column to populate 'Company_Name'
DUNSupdate['Company_Name'] = DUNSupdate['DUNS_Number'].apply(generate_company_name)



# Create a sidebar selectbox for navigation
page = st.sidebar.selectbox("Choose a page", ["Home", "Collaterals", "Changes", "Late payers"])

if page == "Home":
    st.title("Main Dashboard")

    import subprocess

    def RunningUiPathProject():
        """
        This function runs the UiPath project 'Endur Changes'.
        """
        cmd_str = (
            #r"C:\Users\GIANMARIO1\AppData\Local\Programs\UiPath\Studio\UiRobot.exe execute "
            r"C:\Users\GIANMARIO1\AppData\Local\Programs\UiPath\Studio\UiPath.Studio.exe execute "
            #r"--file 'C:\Users\GIANMARIO1\OneDrive - OMV Group\Documents\UiPath\Endur Changes\project.json' "
            #r"--entry 'Main.xaml'"
        )
        subprocess.run(cmd_str, shell=True)


    # Place buttons side by side
    col1, col2 , col3, col4, col5 = st.columns(5)

    with col1:
        # Add a button to run the Excel macro
        if st.button("Refresh the data"):
            try:
                # Initialize COM library
                pythoncom.CoInitialize()

                xl = Dispatch("Excel.Application")
                xl.Workbooks.Open(file_path)
                xl.Application.Run("A_Main")
                xl.ActiveWorkbook.Save()
                xl.Quit()

                st.success("Macro 'A_Main' executed successfully! - The data is up to date")
            except Exception as e:
                st.error(f"Error running the macro: {str(e)}")
    with col2:
        # Add a button to create and open an Excel file
        if st.button("Breaches file"):
            try:
                # Initialize COM library
                pythoncom.CoInitialize()

                # Create a new Excel file in memory
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    breachesdf.to_excel(writer, index=False, sheet_name='Sheet1')
                
                # Save the Excel file to a temporary location
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                temp_file.write(output.getvalue())
                temp_file.close()

                # Debugging: Check if the file is created correctly
                st.write("Excel file created and saved to temporary location")

                # Open the Excel file from the temporary location
                xl = Dispatch("Excel.Application")
                xl.Workbooks.Open(temp_file.name)
                xl.Visible = True
                st.success("Excel file created and opened successfully!")
            except Exception as e:
                st.error(f"Error creating or opening the Excel file: {str(e)}")
            finally:
                pythoncom.CoUninitialize()
    with col3:
        # Add a button to create and open an Excel file
        if st.button("IT Errors"):
            try:
                # Initialize COM library
                pythoncom.CoInitialize()

                # Create a new Excel file in memory
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    ITErrors.to_excel(writer, index=False, sheet_name='Sheet1')
                
                # Save the Excel file to a temporary location
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                temp_file.write(output.getvalue())
                temp_file.close()

                # Debugging: Check if the file is created correctly
                st.write("Excel file created and saved to temporary location")

                # Open the Excel file from the temporary location
                xl = Dispatch("Excel.Application")
                xl.Workbooks.Open(temp_file.name)
                xl.Visible = True
                st.success("Excel file created and opened successfully!")
            except Exception as e:
                st.error(f"Error creating or opening the Excel file: {str(e)}")
            finally:
                pythoncom.CoUninitialize()
    
    with col4:
        if st.button("Open UiPath"):
            try:
                RunningUiPathProject()

            except Exception as e:
                st.error(f"Error Opening Ui Path: {str(e)}")

    with col5:
        if st.button("DnB_Update"):
            try:
                # Initialize COM library
                pythoncom.CoInitialize()

                # Create a new Excel file in memory
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    DUNSupdate.to_excel(writer, index=False, sheet_name='Sheet1')
                
                # Save the Excel file to a temporary location
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                temp_file.write(output.getvalue())
                temp_file.close()

                # Debugging: Check if the file is created correctly
                st.write("Excel file created and saved to temporary location")

                # Open the Excel file from the temporary location
                xl = Dispatch("Excel.Application")
                xl.Workbooks.Open(temp_file.name)
                xl.Visible = True
                st.success("Excel file created and opened successfully!")
            #except Exception as e:
                st.error(f"Error creating or opening the Excel file: {str(e)}")
            finally:
                pythoncom.CoUninitialize()
    
    # Display the relevant fields using AgGrid
    df = df.drop(columns=["country"])
    AgGrid(df)

    # ---- SIDEBAR ----
    st.sidebar.header("Filter the data:")
    selected_values1 = st.sidebar.multiselect("Select the Endur ID", options=df["customer"], default=[])
    selected_values2 = st.sidebar.multiselect("Select the CP Name", options=df["counterparty"], default=[])
    selected_values3 = st.sidebar.multiselect("Select the DUNS No", options=df["duns_number"], default=[])
    selected_values4 = st.sidebar.multiselect("Select the HQ-DUNS No", options=df["hq_duns_number"], default=[])

    AgGrid(df[df['customer'].isin(selected_values1) |
                     df["counterparty"].isin(selected_values2) |
                     df["duns_number"].isin(selected_values3) |
                     df["hq_duns_number"].isin(selected_values4)])
    

                
                
elif page == "Collaterals":
    # Code for the second page goes here
    st.title("Collaterals")
    excel_dir = "//somvat002003/OGP/OGP/SMT/PC/00_DFG-C/03_NEW OGMT CREDIT/06_Collaterals"
    excel_files = glob.glob(os.path.join(excel_dir, "*.xlsx"))
    excel_files.sort(key=os.path.getmtime, reverse=True)
    if excel_files:
        latest_excel_file = excel_files[0]
        st.write(f"Latest Excel file: {latest_excel_file}")
        latest_df = pd.read_excel(latest_excel_file)
        AgGrid(latest_df)

        # ---- SIDEBAR ----
    st.sidebar.header("Filter the data:")
    selected_values11 = st.sidebar.multiselect("Select the Guarantee Receiver", options=latest_df["long_name"], default=[])
    selected_values21 = st.sidebar.multiselect("Select the Guarantee Provider", options=latest_df["guarantor"], default=[])

    latest_df = AgGrid(latest_df[latest_df["long_name"].isin(selected_values11) |
                     latest_df["guarantor"].isin(selected_values21)])

elif page == "Changes":

    import subprocess

    def RunningUiPathProject():
        """
        This function runs the UiPath project 'Endur Changes'.
        """
        cmd_str = (
            #r"C:\Users\GIANMARIO1\AppData\Local\Programs\UiPath\Studio\UiRobot.exe execute "
            r"C:\Users\GIANMARIO1\AppData\Local\Programs\UiPath\Studio\UiPath.Studio.exe execute "
            #r"--file 'C:\Users\GIANMARIO1\OneDrive - OMV Group\Documents\UiPath\Endur Changes\project.json' "
            #r"--entry 'Main.xaml'"
        )
        subprocess.run(cmd_str, shell=True)


    # Code for the third page goes here
    st.title("Changes")
    "NOTES : (i) Open UiPath - The changes are performed via 'Endur Changes', (ii) If something is wrong, then you can open the Endur changes excel file and make the necessary changes"
    Changes_file_path = "//somvat002003/OGP/OGP/SMT/PC/00_DFG-C/03_NEW OGMT CREDIT/04_Changes/Endur Changes/Endur_changes_automated.xlsm"
    Output_df = pd.read_excel(Changes_file_path, sheet_name="Output", usecols="A:M")
    Email_df = pd.read_excel(Changes_file_path, sheet_name="Insurance_Update", usecols="A:E")
    
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if st.button("Export Changes"):
            try:
                # Initialize COM library
                pythoncom.CoInitialize()

                xl = Dispatch("Excel.Application")
                xl.Workbooks.Open(Changes_file_path)
                xl.Application.Run("Export")
                xl.ActiveWorkbook.Save()
                xl.Quit()
                st.success("Succesfully exported!")
            except Exception as e:
                st.error(f"Error running the macro: {str(e)}")

    with col2:
        if st.button("Update Collaterals (email)"):
            try:
                # Initialize COM library
                pythoncom.CoInitialize()

                xl = Dispatch("Excel.Application")
                xl.Workbooks.Open(Changes_file_path)
                xl.Application.Run("SendCollateralsEmail")
                xl.ActiveWorkbook.Save()
                xl.Quit()
            except Exception as e:
                st.error(f"Error running the macro: {str(e)}")

    with col3:
        if st.button("Open excel"):
            try:
                # Initialize COM library
                pythoncom.CoInitialize()

                xl = Dispatch("Excel.Application")
                xl.Workbooks.Open(Changes_file_path)
                xl.ActiveWorkbook.Save()
                st.success("Changes file opened successfully!")
            except Exception as e:
                st.error(f"Error running the macro: {str(e)}")

    with col4:
        if st.button("Open UiPath"):
            try:
                RunningUiPathProject()

            except Exception as e:
                st.error(f"Error Opening Ui Path: {str(e)}")

    

    AgGrid(Output_df)
    AgGrid(Email_df)

elif page == "Late payers":
    print()
