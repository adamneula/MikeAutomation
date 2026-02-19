from tqdm import tqdm
tqdm.pandas()
from Rep_Objects import *
from Pivot_Table import *
from New_And_Additions import *
from SF_Upload_Div_Model import *

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
        
        choice = input("Select an option: ").strip().upper()
        
        if choice == 'Q':
            print("Closing application. Have a good one!")
            break
        
        fitlist = input('Enter FULL PATH of the fit list (<MONTH>-<YEAR): ').strip().replace('"', '')
        fitlist_sheet = input_with_default("Enter Fit List sheet name", "FIT")        
        thisMonth = input('Enter FULL PATH of the Primerica excel file (ModelProvider_AUM_RNC_<MONTH><YEAR>.xlsx): ').strip().replace('"', '')
        thisMonthSheet = input_with_default("Enter current Primerica sheet name", "Account-Rep Details")
        lastMonth = input("Enter FULL PATH of last month's pivot table excel file (ModelProvider_AUM_RNC_<MONTH><YEAR>_Pivot.xlsx): ").strip().replace('"', '')
        
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
            SF_Upload_Sheet(path, 'All Models')
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