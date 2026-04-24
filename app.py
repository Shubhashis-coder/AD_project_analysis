import streamlit as st
import pandas as pd
import re
import io
import plotly.express as px
import numpy as np

st.set_page_config(page_title="Biochar-AD Modelling (Rigorous)", layout="wide")
st.title("🧪 Biochar-Enhanced Anaerobic Digestion Modelling")
st.markdown("**Fully data‑driven RYIELD fractions using O/C ratio correlation**")

# ----------------------------------------------------------------------
# Helper functions
# ----------------------------------------------------------------------
def get_process_param(process_df, param_name):
    val = process_df.loc[process_df['Parameter'] == param_name, 'Value']
    if val.empty:
        raise KeyError(f"Parameter '{param_name}' not found")
    return float(val.values[0])

def buswell_theoretical_ch4(ultimate_df):
    fw = ultimate_df[ultimate_df['Material'].str.lower() == 'foodwaste']
    if fw.empty:
        raise ValueError("FoodWaste not found in Ultimate_Analysis")
    C = fw['C_wt%'].values[0] / 100
    H = fw['H_wt%'].values[0] / 100
    O = fw['O_wt%'].values[0] / 100
    N = fw['N_wt%'].values[0] / 100
    mol_C = C / 12.01
    mol_H = H / 1.008
    mol_O = O / 16.00
    mol_N = N / 14.01
    a = mol_H / mol_C
    b = mol_O / mol_C
    c = mol_N / mol_C
    nu_CH4 = 0.5 + a/8 - b/4 - 3*c/8
    M_unit = 12.01 + a*1.008 + b*16.00 + c*14.01
    mol_CH4_per_gVS = nu_CH4 / M_unit
    return mol_CH4_per_gVS * 22400   # mL/g VS

def acetate_carbon_fraction_from_OC(ultimate_df):
    """
    Compute the fraction of total carbon that goes to acetate during acidogenesis
    based on the atomic O/C ratio. Literature correlation (Labatut et al., 2011).
    """
    fw = ultimate_df[ultimate_df['Material'].str.lower() == 'foodwaste']
    if fw.empty:
        raise ValueError("FoodWaste not found")
    C_wt = fw['C_wt%'].values[0]
    O_wt = fw['O_wt%'].values[0]
    # Atomic O/C ratio
    O_C_atomic = (O_wt / 16.00) / (C_wt / 12.01)
    # Empirical equation: acetate carbon fraction = 0.20 + 0.40*exp(-0.5*O/C)
    frac = 0.20 + 0.40 * np.exp(-0.5 * O_C_atomic)
    # Clamp to reasonable range (0.35 - 0.65)
    return max(0.35, min(0.65, frac))

def ryield_fractions_thermo(ultimate_df, theoretical_mL_per_gVS, acetate_C_frac):
    """
    RYIELD fractions using carbon‑based acetate yield + electron balance.
    All fractions in g/g VS.
    """
    fw = ultimate_df[ultimate_df['Material'].str.lower() == 'foodwaste']
    if fw.empty:
        raise ValueError("FoodWaste not found")
    
    C_frac = fw['C_wt%'].values[0] / 100
    Ash = fw['Ash_wt%'].values[0] / 100 if 'Ash_wt%' in fw else 0
    total_C_mol = C_frac / 12.01
    
    # Carbon to acetate
    C_to_acetate_mol = total_C_mol * acetate_C_frac
    mol_acetate = C_to_acetate_mol / 2
    mass_acetate = mol_acetate * 59.0
    
    # Remaining carbon -> CO2
    C_to_CO2_mol = total_C_mol - C_to_acetate_mol
    mass_CO2 = C_to_CO2_mol * 44.0
    
    # Total electrons from COD
    COD = theoretical_mL_per_gVS / 350.0
    total_e = COD * 4.0   # meq/g VS
    e_acetate = mol_acetate * 8.0
    e_H2 = max(0.0, total_e - e_acetate)
    mol_H2 = e_H2 / 2.0
    mass_H2 = mol_H2 * 2.0
    
    # Digest from ash
    if Ash < 1.0:
        ash_vs = Ash / (1 - Ash) if Ash < 1 else 0.2
    else:
        ash_vs = 0.2
    mass_digest = min(0.30, max(0.10, ash_vs + 0.05))
    
    # Water closure
    sum_others = mass_acetate + mass_H2 + mass_CO2 + mass_digest
    mass_water = max(0.0, 1.0 - sum_others)
    
    # Normalise
    total = mass_acetate + mass_H2 + mass_CO2 + mass_water + mass_digest
    if total > 0:
        mass_acetate /= total
        mass_H2 /= total
        mass_CO2 /= total
        mass_water /= total
        mass_digest /= total
    
    return pd.DataFrame({
        'Product': ['Acetate', 'H2', 'CO2', 'H2O', 'Digest'],
        'Fraction': [mass_acetate, mass_H2, mass_CO2, mass_water, mass_digest]
    })

# ----------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------
uploaded_file = st.sidebar.file_uploader("📂 Upload AD_Project_Input.xlsx", type=["xlsx"])

if uploaded_file is not None:
    try:
        methane = pd.read_excel(uploaded_file, sheet_name='Daily_Methane')
        process = pd.read_excel(uploaded_file, sheet_name='Process_Inputs')
        ultimate = pd.read_excel(uploaded_file, sheet_name='Ultimate_Analysis')
        st.sidebar.success("✅ File loaded")
        
        # Automatically compute acetate carbon fraction from O/C ratio
        auto_acetate_frac = acetate_carbon_fraction_from_OC(ultimate)
        st.sidebar.markdown(f"**📐 O/C ratio derived acetate C fraction:** `{auto_acetate_frac:.3f}`")
        
        # Optional manual override (for sensitivity)
        use_manual = st.sidebar.checkbox("Override with manual value (sensitivity analysis)", value=False)
        if use_manual:
            manual_frac = st.sidebar.slider("Manual acetate carbon fraction", 0.30, 0.65, auto_acetate_frac, 0.01)
            acetate_C_frac = manual_frac
            st.sidebar.info(f"Using manual: {acetate_C_frac:.3f}")
        else:
            acetate_C_frac = auto_acetate_frac
            st.sidebar.success("Using correlation from O/C ratio")
        
        # ---------- Calculations (unchanged from original) ----------
        VS_added = get_process_param(process, 'VS_added')
        final_row = methane.iloc[-1]
        methane_cols = [col for col in methane.columns if col != 'Day']
        normalized = {}
        for col in methane_cols:
            try:
                normalized[col] = final_row[col] / VS_added
            except:
                normalized[col] = None
        
        control_col = 'CH4_No_Biochar' if 'CH4_No_Biochar' in normalized else 'No_Biochar'
        biochar_cols = [c for c in normalized if c != control_col and normalized[c] is not None]
        opt_col = max(biochar_cols, key=lambda c: normalized[c])
        opt_yield = normalized[opt_col]
        
        dose_match = re.search(r'(\d+(?:\.\d+)?)pct', opt_col, re.IGNORECASE)
        if dose_match:
            opt_dose = float(dose_match.group(1)) / 100.0
        else:
            opt_dose = get_process_param(process, 'Optimum_Biochar_Dose')
        
        y_without = normalized[control_col]
        y_with = opt_yield
        B = opt_dose
        alpha = ((y_with / y_without) - 1) / B
        base_k = get_process_param(process, 'Base_k')
        k_eff = base_k * (1 + alpha * B)
        
        yield_table = pd.DataFrame([
            [col.replace('CH4_','').replace('_Biochar',''), val]
            for col, val in normalized.items() if val is not None
        ], columns=['Case', 'Methane_Yield_mL_CH4_gVS'])
        
        alpha_table = pd.DataFrame({
            'Parameter': ['Y_without', 'Y_with', 'Biochar_Dose', 'Alpha'],
            'Value': [y_without, y_with, B, alpha]
        })
        k_table = pd.DataFrame({
            'Parameter': ['Base_k', 'Effective_k'],
            'Value': [base_k, k_eff]
        })
        
        theoretical_buswell = buswell_theoretical_ch4(ultimate)
        ryield = ryield_fractions_thermo(ultimate, theoretical_buswell, acetate_C_frac)
        
        buswell_table = pd.DataFrame({
            'Parameter': ['Theoretical_CH4_mL_gVS', 'Experimental_optimum_mL_gVS', 'Difference_%', 'Note'],
            'Value': [theoretical_buswell, y_with, f"{(y_with/theoretical_buswell -1)*100:.1f}%",
                      'Negative = experimental lower (kinetic limitation)']
        })
        
        # Output Excel buffer
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            yield_table.to_excel(writer, sheet_name='Methane_Yield', index=False)
            alpha_table.to_excel(writer, sheet_name='Alpha_Calc', index=False)
            k_table.to_excel(writer, sheet_name='Kinetic_k', index=False)
            ryield.to_excel(writer, sheet_name='RYIELD_Table', index=False)
            buswell_table.to_excel(writer, sheet_name='Buswell_Validation', index=False)
        output_buffer.seek(0)
        
        # ---------- Display Results ----------
        st.header("📊 Results & Visualisations")
        col1, col2 = st.columns(2)
        with col1:
            fig_bar = px.bar(yield_table, x='Case', y='Methane_Yield_mL_CH4_gVS', 
                             title="Methane Yield", text='Methane_Yield_mL_CH4_gVS')
            fig_bar.update_traces(textposition='outside')
            st.plotly_chart(fig_bar, use_container_width=True)
        with col2:
            st.metric("Optimum Dose", f"{B*100:.1f}%")
            st.metric("Enhancement α", f"{alpha:.4f}")
            st.metric("Effective k (day⁻¹)", f"{k_eff:.4f}")
        
        if methane.shape[0] > 1:
            melt_df = methane.melt(id_vars=['Day'], var_name='Case', value_name='CH4_mL')
            fig_line = px.line(melt_df, x='Day', y='CH4_mL', color='Case',
                               title="Cumulative Methane Production")
            st.plotly_chart(fig_line, use_container_width=True)
        
        bus_comp = pd.DataFrame({
            'Source': ['Theoretical (Buswell)', 'Experimental (Optimum)'],
            'Yield_mL_gVS': [theoretical_buswell, y_with]
        })
        fig_bus = px.bar(bus_comp, x='Source', y='Yield_mL_gVS', title="Theoretical vs Experimental",
                         text='Yield_mL_gVS')
        fig_bus.update_traces(textposition='outside')
        st.plotly_chart(fig_bus, use_container_width=True)
        
        st.subheader("✅ RYIELD Fractions (from O/C correlation)")
        fig_ryield = px.bar(ryield, x='Product', y='Fraction', title="Food Waste Decomposition",
                            text=ryield['Fraction'].round(3))
        fig_ryield.update_traces(textposition='outside')
        st.plotly_chart(fig_ryield, use_container_width=True)
        
        st.dataframe(ryield)
        
        # Additional information: O/C ratio and derived fraction
        o_c_atomic = (ultimate[ultimate['Material'].str.lower()=='foodwaste']['O_wt%'].values[0]/16.00) / \
                     (ultimate[ultimate['Material'].str.lower()=='foodwaste']['C_wt%'].values[0]/12.01)
        st.info(f"📐 **Atomic O/C ratio:** {o_c_atomic:.3f} → **Acetate carbon fraction:** {acetate_C_frac:.3f} (from empirical correlation)")
        
        st.sidebar.download_button("📥 Download AD_Project_Output.xlsx", data=output_buffer,
                                   file_name="AD_Project_Output.xlsx")
        
    except Exception as e:
        st.error(f"Error: {e}")
        st.stop()
else:
    st.info("👈 Upload your `AD_Project_Input.xlsx` file")