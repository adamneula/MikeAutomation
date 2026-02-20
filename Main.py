from tqdm import tqdm
tqdm.pandas()
from Rep_Objects import *
from Pivot_Table import *
from New_And_Additions import *
from SF_Upload_Div_Model import *
import os

def main():
    while True:
        print("\n" + "="*40)
        print("      GENTER CAPITAL AUTOMATION")
        print("=" * 40)
        print("1. Generate Primerica AUM Pivot Table")
        print("2. Run Primerica Div Model Additions + Opens")
        print("3. Run GenT and GenM Additions + Opens")
        print("4. Run All Pipelines")
        print("Q. Quit")
        print("-" * 40)
        print("So long as the fit list, Account-Rep details, and prior month's pivot tables are stored in this directory with default names,")
        print("you should be able to just hit enter through all the input screens with the automatic value shown in the square brackets.")
        print("-" * 40)
        
        choice = input("Select an option: ").strip().upper()
        
        if choice == 'Q':
            print("Closing application. Have a good one!")
            break
        
        base_path = os.path.dirname(os.path.abspath(__file__))
        
        # Calculate offsets
        prev_month = pd.Timestamp.now() - pd.DateOffset(months=1)
        two_months_ago = pd.Timestamp.now() - pd.DateOffset(months=2)
        
        # Format the strings
        fitlist_date = f"{prev_month.month}-{prev_month.strftime('%y')}"  # e.g., "1-26" or "12-25"
        thisMonth_date = prev_month.strftime('%b%Y').upper()              # e.g., "JAN2026"
        lastMonth_date = two_months_ago.strftime('%b%Y').upper()          # e.g., "DEC2025"
        
        # Build the full default paths
        fitlist_default = os.path.join(base_path, f"{fitlist_date}.xlsx")
        thisMonth_default = os.path.join(base_path, f"ModelProvider_AUM_RNC_{thisMonth_date}.xlsx")
        lastMonth_default = os.path.join(base_path, f"ModelProvider_AUM_RNC_{lastMonth_date}_Pivot.xlsx")
        
        fitlist = input_with_default('Enter FULL PATH of the fit list', fitlist_default).strip().replace('"', '')
        fitlist_sheet = input_with_default("Enter Fit List sheet name", "FIT")        
        thisMonth = input_with_default('Enter FULL PATH of the Primerica excel file', thisMonth_default).strip().replace('"', '')
        thisMonthSheet = input_with_default("Enter current Primerica sheet name", "Account-Rep Details")
        lastMonth = input_with_default("Enter FULL PATH of last month's pivot table excel file", lastMonth_default).strip().replace('"', '')
        
        if choice == '1':
            prior_month_str = (pd.Timestamp.now() - pd.DateOffset(months=1)).strftime('%b %y')
            suggested_sheet = f"AUM Pivot - {prior_month_str}"
            
            lastMonthTableSheet = input_with_default('Enter the name of the pivot table sheet', suggested_sheet)

            load_reps_from_xlsx(fitlist, fitlist_sheet)
            attribute_accounts(thisMonth, thisMonthSheet)
            load_previous_month_data(lastMonth, lastMonthTableSheet)
            export_to_pivot(fitlist, fitlist_sheet, thisMonth, thisMonthSheet, lastMonth, lastMonthTableSheet)
        elif choice == '2':
            lastMonthAccountSheet = input_with_default("Enter the name of the sheet on last month's Primerica table's file",  "Account-Rep Details")
            
            load_reps_from_xlsx(fitlist, fitlist_sheet)
            path = Primerica_Div_Model_New_And_Addition(thisMonth, thisMonthSheet, lastMonth, lastMonthAccountSheet)
            SF_Upload_Sheet(path, 'Primerica Div Model')
        elif choice == '3':
            lastMonthAccountSheet = input_with_default("Enter the name of the sheet on last month's Primerica table's file",  "Account-Rep Details")
            
            load_reps_from_xlsx(fitlist, fitlist_sheet)
            path = GenT_GenM_New_And_Addition(thisMonth, thisMonthSheet, lastMonth, lastMonthAccountSheet)
            SF_Upload_Sheet(path, 'GENT and GENM')
        elif choice == '4':
            prior_month_str = (pd.Timestamp.now() - pd.DateOffset(months=2)).strftime('%b %y')
            suggested_sheet = f"AUM Pivot - {prior_month_str}"
            
            lastMonthTableSheet = input_with_default('Enter the name of the pivot table sheet', suggested_sheet)
            lastMonthAccountSheet = input_with_default("Enter the name of the sheet on last month's Primerica table's file",  "Account-Rep Details")
            
            load_reps_from_xlsx(fitlist, fitlist_sheet)
            attribute_accounts(thisMonth, thisMonthSheet)
            load_previous_month_data(lastMonth, lastMonthTableSheet)
            export_to_pivot(fitlist, fitlist_sheet, thisMonth, thisMonthSheet, lastMonth, lastMonthTableSheet)
            path = Primerica_Div_Model_New_And_Addition(thisMonth, thisMonthSheet, lastMonth, lastMonthAccountSheet)
            SF_Upload_Sheet(path, 'Primerica Div Model')
            path = GenT_GenM_New_And_Addition(thisMonth, thisMonthSheet, lastMonth, lastMonthAccountSheet)
            SF_Upload_Sheet(path, 'GENT and GENM')
        else:
            print("Invalid selection. Please enter 1, 2, 3, 4, or Q.")

if __name__ == "__main__":
    main()  