import pandas as pd

input_file = 'AD_Project_Input.xlsx'
output_file = 'AD_Project_Output.xlsx'

methane = pd.read_excel(input_file, sheet_name='Daily_Methane')
process = pd.read_excel(input_file, sheet_name='Process_Inputs')

final_row = methane.iloc[-1]

yield_table = pd.DataFrame({
    'Case': ['No Biochar','2% Biochar','5% Biochar','8% Biochar','10% Biochar'],
    'Methane_Yield_mL_CH4_gVS': [
        final_row['CH4_No_Biochar'],
        final_row['CH4_2pct'],
        final_row['CH4_5pct_Biochar'],
        final_row['CH4_8pct'],
        final_row['CH4_10pct']
    ]
})

y_without = final_row['CH4_No_Biochar']
y_with = final_row['CH4_5pct_Biochar']
B = float(process.loc[process['Parameter']=='Optimum_Biochar_Dose','Value'].values[0])

alpha = ((y_with / y_without) - 1) / B

alpha_table = pd.DataFrame({
    'Parameter': ['Y_without','Y_with','Biochar_Dose','Alpha'],
    'Value': [y_without, y_with, B, alpha]
})

base_k = float(process.loc[process['Parameter']=='Base_k','Value'].values[0])
k_eff = base_k * (1 + alpha * B)

k_table = pd.DataFrame({
    'Parameter': ['Base_k','Effective_k'],
    'Value': [base_k, k_eff]
})

ryield = pd.DataFrame({
    'Product': ['Acetate','H2','CO2','H2O','Digest'],
    'Fraction': [0.44,0.07,0.19,0.15,0.15]
})

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    yield_table.to_excel(writer, sheet_name='Methane_Yield', index=False)
    alpha_table.to_excel(writer, sheet_name='Alpha_Calc', index=False)
    k_table.to_excel(writer, sheet_name='Kinetic_k', index=False)
    ryield.to_excel(writer, sheet_name='RYIELD_Table', index=False)

print('Done')
