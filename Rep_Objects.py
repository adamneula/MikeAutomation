import pandas as pd

reps = {}
IDtoName = {}
# States = {
#     'AL', 'AK', 'AZ', 'AR', 'AB', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 
#     'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD', 
#     'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 
#     'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'PR', 'RI', 'SC', 
#     'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY', 'DC'
# }

class Representatives():
    def __init__(self, name: str, ID: str, state: str, email: str, AE: str, territory: str, total: float):
        self.Advisor_Name = name
        self.Primary_Rep_ID = ID
        self.True_State = state
        self.Email = email
        self.AE = AE
        self.Territory = territory
        self.Ranking = None
        self.Sum_of_Total_Assets = 0
        self.Previous_Month_AUM = None
        self.MoM_Change = None
        self.Dollar_Val_Change = None
        self.Lifetime_Total = total
        
    def __hash__(self):
        return hash(self.Primary_Rep_ID)

    def __eq__(self, other):
        if not isinstance(other, Representatives):
            return False
        return self.Primary_Rep_ID == other.Primary_Rep_ID
    
    def __str__(self):
        return f'{self.Advisor_Name} {self.True_State} {self.Primary_Rep_ID} {self.Email} {self.Territory} AE: {self.AE} Balance: {self.Sum_of_Total_Assets}'

    def add_account(self, amount):
        self.Sum_of_Total_Assets += amount
        
def load_reps_from_xlsx(Fit_List_Dir, Fit_List_Sheet_Name):
    global reps
    
    # header=1 skips the 'Owned...' row and uses the 'Code, Mutual...' row as headers
    df = pd.read_excel(Fit_List_Dir, sheet_name=Fit_List_Sheet_Name, header=1)
    
    # Standardize column names to remove any accidental spaces
    df.columns = df.columns.str.strip()
    
    # Drop rows where ID is missing
    df = df.dropna(subset=['ID'])
    
    for _, row in df.iterrows():
        first_name = str(row['First']).strip()
        last_name = str(row['Last']).strip()
        if first_name.lower() == 'christophe': first_name = 'CHRISTOPHER'
        elif first_name.lower() == 'theodore' and last_name.lower() == 'lund': first_name = 'TED'
        elif first_name.lower() == 'danny' and last_name.lower() == 'creswell': first_name = 'DANIEL'
        full_name = f"{first_name} {last_name}"
        clean_ID = str(row['ID']).replace(' ', '').strip()        
        IDtoName[clean_ID] = full_name.lower()
        
        total = float(row['LifeTime'])
        if full_name.lower() in reps:
            if reps[full_name.lower()].Lifetime_Total > total: continue
        state = str(row['State']).strip()
        email = str(row['Pol Email']).strip()
        territory = str(row['Territory']).strip()
        AE = ''
        #Sets central to East region and assigns AE accordingly
        if state.lower() in ['ok', 'ks']:
            territory = 'East'
            AE = 'Rob Hunt'
        if territory.lower() == 'central':
            territory = 'East' 
            AE = 'Rob Hunt'
        elif territory == 'East': AE = 'Rob Hunt'
        elif territory == 'West': AE = 'MeiWah Wong'
        reps[full_name.lower()] = Representatives(full_name, clean_ID, state, email, AE, territory, total)
        
def rep_lookup(input_str) -> Representatives:
    global reps, IDtoName
    
    if not input_str or str(input_str).lower() == 'nan':
        return None

    input_clean = str(input_str).strip().upper()
    resolved_name_lower = IDtoName.get(input_clean)
    
    if not resolved_name_lower:
        resolved_name_lower = IDtoName.get(input_clean.replace(" ", ""))
    if not resolved_name_lower:
        resolved_name_lower = IDtoName.get(input_clean[:5])
    target_name_lower = resolved_name_lower if resolved_name_lower else input_clean.lower()

    parts = target_name_lower.split()
    if parts:
        first = parts[0]
        last = " ".join(parts[1:])
        
        if first == 'christophe':
            target_name_lower = " ".join(['christopher', last])
        elif first == 'danny' and last == 'creswell':
            target_name_lower = 'daniel creswell'
        elif first == 'theodore' and last == 'lund':
            target_name_lower = 'ted lund'

    return reps.get(target_name_lower)

def get_reps():
    return reps