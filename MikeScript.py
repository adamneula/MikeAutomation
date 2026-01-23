import pandas as pd
import numpy as np
import os

# First edit on corp machine
# Global dictionary to store objects
reps = {}
regionMap = {}
States = {
    'AL', 'AK', 'AZ', 'AR', 'AB', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 
    'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD', 
    'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 
    'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'PR', 'RI', 'SC', 
    'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY', 'DC'
}

class Representatives():
    def __init__(self, name: str, ID: str, state: str, email: str, AE: str, territory: str):
        self.Advisor_Name = name
        self.Primary_Rep_ID = ID
        self.True_State = state
        self.Email = email
        self.AE = AE
        self.Territory = territory
        self.Ranking = None
        self.Sum_of_Total_Assets = None
        self.Previous_Month_AUM = None
        self.MoM_Change = None
        self.Dollar_Val_Change = None
        
    def __hash__(self):
        return hash(self.Primary_Rep_ID)

    def __eq__(self, other):
        if not isinstance(other, Representatives):
            return False
        return self.Primary_Rep_ID == other.Primary_Rep_ID
    
    def __str__(self):
        return f'{self.Advisor_Name} {self.True_State} {self.Primary_Rep_ID} {self.Email} {self.Territory} AE: {self.AE}'

def load_reps_from_csv():
    global reps
    
    # Setup pathing
    base_dir = os.path.dirname(os.path.abspath(__file__))
    #sheet_name = input('Enter sheet name (it should be stored in the same directory as this script. Give the name and filetype, such as 12-25.xlsx): ')
    file_path = os.path.join(base_dir, '12-25.xlsx')
    
    # header=1 skips the 'Owned...' row and uses the 'Code, Mutual...' row as headers
    df = pd.read_excel(file_path, sheet_name='FIT', header=1)
    
    # Standardize column names to remove any accidental spaces
    df.columns = df.columns.str.strip()
    
    # Drop rows where ID is missing
    df = df.dropna(subset=['ID'])
    
    for _, row in df.iterrows():
        full_name = f"{row['First']} {row['Last']}"
        clean_ID = str(row['ID']).replace(' ', '').strip()
        state = str(row['State']).strip()
        email = str(row['Pol Email']).strip()
        territory = str(row['Territory']).strip()
        AE = ""
        #Sets central to East region
        if territory == 'Central':
            territory = 'East'
            AE = 'Rob Hunt'
        elif territory == 'East':
            AE = 'Rob Hunt'
        elif territory == 'West':
            AE = 'MeiWah Wong'
            
        reps[clean_ID] = Representatives(full_name, clean_ID, state, email, AE, territory)


load_reps_from_csv()
print(f'loaded {len(reps)} reps')
print(reps['60916'])
