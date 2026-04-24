import pandas as pd
import re

# ----------------------------------------------------------------------
# File names
# ----------------------------------------------------------------------
input_file = 'AD_Project_Input.xlsx'
output_file = 'AD_Project_Output.xlsx'

# ----------------------------------------------------------------------
# Helper function to read a parameter from Process_Inputs
# ----------------------------------------------------------------------
def get_process_param(process_df, param_name):
    val = process_df.loc[process_df['Parameter'] == param_name, 'Value']
    if val.empty:
        raise KeyError(f"Parameter '{param_name}' not found in Process_Inputs sheet")
    return float(val.values[0])

# ----------------------------------------------------------------------
# Buswell equation: theoretical methane potential (mL/g VS)
# ----------------------------------------------------------------------
def buswell_theoretical_ch4(ultimate_df):
    """
    Calculate theoretical methane potential (mL CH4/g VS) at STP
    using the Buswell equation from ultimate analysis of FoodWaste.
    """
    fw = ultimate_df[ultimate_df['Material'].str.lower() == 'foodwaste']
    if fw.empty:
        raise ValueError("FoodWaste not found in Ultimate_Analysis sheet")
    
    # Extract mass fractions (as decimals)
    C = fw['C_wt%'].values[0] / 100
    H = fw['H_wt%'].values[0] / 100
    O = fw['O_wt%'].values[0] / 100
    N = fw['N_wt%'].values[0] / 100
    
    # Moles per gram of VS
    mol_C = C / 12.01
    mol_H = H / 1.008
    mol_O = O / 16.00
    mol_N = N / 14.01
    
    # Empirical formula per mole of carbon
    a = mol_H / mol_C   # H/C
    b = mol_O / mol_C   # O/C
    c = mol_N / mol_C   # N/C
    
    # Buswell equation: moles CH4 per mole C
    nu_CH4 = 0.5 + a/8 - b/4 - 3*c/8
    
    # Mass of the C-unit (g per mole of C)
    M_unit = 12.01 + a*1.008 + b*16.00 + c*14.01
    
    # Moles CH4 per gram VS
    mol_CH4_per_gVS = nu_CH4 / M_unit
    
    # Convert to mL/g VS at STP (22.4 L/mol = 22400 mL/mol)
    mL_CH4_per_gVS = mol_CH4_per_gVS * 22400
    return mL_CH4_per_gVS

# ----------------------------------------------------------------------
# Thermodynamic RYIELD fractions (electron balance + carbon balance)
# ----------------------------------------------------------------------
def ryield_fractions_thermo(ultimate_df, theoretical_mL_per_gVS):
    """
    Calculate RYIELD fractions (Acetate, H2, CO2, H2O, Digest)
    using electron balance derived from theoretical CH4 (COD)
    and ash content for digest.
    """
    fw = ultimate_df[ultimate_df['Material'].str.lower() == 'foodwaste']
    if fw.empty:
        raise ValueError("FoodWaste not found for RYIELD calculation")
    
    # Ash fraction (dry basis)
    Ash = fw['Ash_wt%'].values[0] / 100 if 'Ash_wt%' in fw else 0
    
    # ---- Step 1: Total COD from theoretical methane ----
    # 1 g COD = 350 mL CH4 at STP
    COD = theoretical_mL_per_gVS / 350.0   # g O2/g VS
    total_e = COD * 4.0                    # meq electrons/g VS
    
    # ---- Step 2: Electron distribution (typical for acidogenesis) ----
    # Fractions can be adjusted based on literature (e.g., 75% to acetate, 15% to H2, 10% to CO2)
    e_frac_acetate = 0.75
    e_frac_H2 = 0.15
    # Remaining 10% is assumed to go to CO2 (no electron transfer)
    
    e_acetate = total_e * e_frac_acetate
    e_H2 = total_e * e_frac_H2
    
    # ---- Step 3: Convert electrons to masses ----
    # Acetate (CH3COO-): 8 e- per mole, molar mass = 59 g/mol
    mol_acetate = e_acetate / 8.0
    mass_acetate = mol_acetate * 59.0          # g/g VS
    
    # H2: 2 e- per mole, molar mass = 2 g/mol
    mol_H2 = e_H2 / 2.0
    mass_H2 = mol_H2 * 2.0                    # g/g VS
    
    # ---- Step 4: Carbon balance to get CO2 ----
    # Total carbon in 1 g VS
    C_mass = fw['C_wt%'].values[0] / 100
    total_C_mol = C_mass / 12.01
    # Carbon in acetate (2 C per molecule)
    C_in_acetate = mol_acetate * 2
    # Remaining carbon goes to CO2
    C_in_CO2 = total_C_mol - C_in_acetate
    mass_CO2 = C_in_CO2 * 44.0                # g/g VS
    
    # ---- Step 5: Digest fraction from ash ----
    if Ash < 1.0:
        ash_vs = Ash / (1 - Ash) if Ash < 1 else 0.2
    else:
        ash_vs = 0.2
    digest_raw = min(0.30, max(0.10, ash_vs + 0.05))
    
    # ---- Step 6: Water fraction (balance to reach total mass = 1) ----
    # Sum of organic products + digest
    sum_organic_digest = mass_acetate + mass_H2 + mass_CO2 + digest_raw
    # Water makes up the rest (ensure it is not negative)
    water_raw = max(0.05, 1.0 - sum_organic_digest)
    
    # ---- Step 7: Normalise to sum exactly 1 (avoid rounding errors) ----
    total_raw = mass_acetate + mass_H2 + mass_CO2 + water_raw + digest_raw
    frac_acetate = mass_acetate / total_raw
    frac_H2 = mass_H2 / total_raw
    frac_CO2 = mass_CO2 / total_raw
    frac_water = water_raw / total_raw
    frac_digest = digest_raw / total_raw
    
    # Final safety normalisation
    total = frac_acetate + frac_H2 + frac_CO2 + frac_water + frac_digest
    if abs(total - 1.0) > 1e-6:
        frac_acetate /= total
        frac_H2 /= total
        frac_CO2 /= total
        frac_water /= total
        frac_digest /= total
    
    return pd.DataFrame({
        'Product': ['Acetate', 'H2', 'CO2', 'H2O', 'Digest'],
        'Fraction': [frac_acetate, frac_H2, frac_CO2, frac_water, frac_digest]
    })

# ----------------------------------------------------------------------
# Main script
# ----------------------------------------------------------------------
def main():
    # 1. Read all input sheets
    try:
        methane = pd.read_excel(input_file, sheet_name='Daily_Methane')
        process = pd.read_excel(input_file, sheet_name='Process_Inputs')
        ultimate = pd.read_excel(input_file, sheet_name='Ultimate_Analysis')
    except Exception as e:
        raise FileNotFoundError(f"Error reading input file/sheets: {e}")
    
    # 2. Get VS_added and normalise methane yields
    VS_added = get_process_param(process, 'VS_added')
    final_row = methane.iloc[-1]
    methane_cols = [col for col in methane.columns if col != 'Day']
    
    normalized = {}
    for col in methane_cols:
        try:
            normalized[col] = final_row[col] / VS_added
        except:
            print(f"Warning: Could not normalise column {col}")
            normalized[col] = None
    
    # 3. Automatic optimum biochar dose detection
    # Identify control column
    control_col = None
    for possible in ['CH4_No_Biochar', 'No_Biochar']:
        if possible in normalized:
            control_col = possible
            break
    if control_col is None:
        raise KeyError("Control column (No Biochar) not found in Daily_Methane")
    
    # Biochar columns
    biochar_cols = [c for c in normalized.keys() if c != control_col and normalized[c] is not None]
    if not biochar_cols:
        raise ValueError("No biochar columns found in Daily_Methane")
    
    opt_col = max(biochar_cols, key=lambda c: normalized[c])
    opt_yield = normalized[opt_col]
    
    # Extract biochar dose from column name (e.g., 'CH4_5pct_Biochar' -> 0.05)
    dose_match = re.search(r'(\d+(?:\.\d+)?)pct', opt_col, re.IGNORECASE)
    if dose_match:
        opt_dose = float(dose_match.group(1)) / 100.0
    else:
        # Fallback to Process_Inputs
        opt_dose = get_process_param(process, 'Optimum_Biochar_Dose')
        print(f"Could not parse dose from column name, using Process_Inputs: {opt_dose}")
    
    y_without = normalized[control_col]
    y_with = opt_yield
    B = opt_dose
    
    # 4. Methane Yield table
    yield_data = []
    for col, val in normalized.items():
        if val is not None:
            case_name = col.replace('CH4_', '').replace('_Biochar', '')
            yield_data.append([case_name, val])
    yield_table = pd.DataFrame(yield_data, columns=['Case', 'Methane_Yield_mL_CH4_gVS'])
    
    # 5. Alpha calculation
    alpha = ((y_with / y_without) - 1) / B
    alpha_table = pd.DataFrame({
        'Parameter': ['Y_without', 'Y_with', 'Biochar_Dose', 'Alpha'],
        'Value': [y_without, y_with, B, alpha]
    })
    
    # 6. Effective kinetic constant
    base_k = get_process_param(process, 'Base_k')
    k_eff = base_k * (1 + alpha * B)
    k_table = pd.DataFrame({
        'Parameter': ['Base_k', 'Effective_k'],
        'Value': [base_k, k_eff]
    })
    
    # 7. Buswell theoretical methane & RYIELD fractions (thermodynamic)
    theoretical_buswell = buswell_theoretical_ch4(ultimate)
    ryield = ryield_fractions_thermo(ultimate, theoretical_buswell)
    
    # 8. Buswell validation sheet
    buswell_table = pd.DataFrame({
        'Parameter': [
            'Theoretical_CH4_mL_gVS (Buswell)',
            'Experimental_optimum_CH4_mL_gVS',
            'Difference_%',
            'Note'
        ],
        'Value': [
            theoretical_buswell,
            y_with,
            f"{(y_with / theoretical_buswell - 1) * 100:.1f}%",
            'Positive means experimental > theoretical (kinetic enhancement)'
        ]
    })
    
    # 9. Write all output sheets
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        yield_table.to_excel(writer, sheet_name='Methane_Yield', index=False)
        alpha_table.to_excel(writer, sheet_name='Alpha_Calc', index=False)
        k_table.to_excel(writer, sheet_name='Kinetic_k', index=False)
        ryield.to_excel(writer, sheet_name='RYIELD_Table', index=False)
        buswell_table.to_excel(writer, sheet_name='Buswell_Validation', index=False)
    
    print(f"Done! Results saved to {output_file}")
    print(f"Optimum biochar dose: {B*100:.1f}% (column: {opt_col}, yield: {y_with:.2f} mL/g VS)")
    print(f"Alpha = {alpha:.4f}, k_eff = {k_eff:.4f} day⁻¹")

# ----------------------------------------------------------------------
if __name__ == "__main__":
    main()