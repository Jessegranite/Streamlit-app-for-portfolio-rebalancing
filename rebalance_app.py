import pandas as pd
import os
import glob
import streamlit as st
import matplotlib.pyplot as plt
import openpyxl

# === CONFIG ===
st.set_page_config(page_title="Portfolio Rebalancer", layout="centered")

st.markdown("""
    <style>
        body {
            background-color: #f4f1ee;
            color: #3e3c3d;
            font-family: 'Inter', sans-serif;
        }
        .stButton>button, .stDownloadButton>button {
            background-color: #ada69e;
            color: white;
            border-radius: 6px;
            font-weight: 500;
        }
        .stDataFrame thead {
            background-color: #dcd6cf;
            color: black;
        }
        .control-row {
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 0.5rem 0;
        }
        .lock-wrap {
            display: flex;
            align-items: center;
            gap: 5px;
            justify-content: flex-end;
        }
        .lock-icon {
            font-size: 1rem;
            color: #f41ef1;
        }
        .block-label {
            font-weight: 500;
            font-size: 0.95rem;
            text-transform: uppercase;
            padding-top: 0.45rem;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Pierre's Portfolio Rebalancer")

# === Load most recent file ===
# === Upload Excel file ===
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file is None:
    st.warning("Please upload an Excel file to continue.")
    st.stop()

# Save temporarily for processing
with open("uploaded_file.xlsx", "wb") as f:
    f.write(uploaded_file.read())
latest_file = "uploaded_file.xlsx"


# === Extract client name from A3 ===
wb = openpyxl.load_workbook(latest_file)
ws = wb.active
client_name = ws["A3"].value or "Client"
st.markdown(f"""
    <div style='background-color: #eae6e1; padding: 1rem; border-radius: 10px; text-align: center;
    font-size: 1.6rem; font-weight: bold; color: #3e3c3d; border: 1px solid #ccc; margin-top: 1rem; margin-bottom: 1.5rem;'>
        👤 {client_name}
    </div>
""", unsafe_allow_html=True)

# === Load and clean data ===
df = pd.read_excel(latest_file)
df.drop(df.columns[0], axis=1, inplace=True)
df.rename(columns={
    df.columns[0]: "Asset Class",
    df.columns[1]: "Security Name",
    df.columns[2]: "Quantity",
    df.columns[3]: "Market Value (CAD)"
}, inplace=True)

df["Asset Class"] = df["Asset Class"].ffill().str.strip()
df = df[df["Quantity"].notna()]

# === Define desired order ===
desired_order = ["Cash & Cash Equivalents", "Bonds", "Canadian Equity", "Global Equity"]

# === Group current totals ===
grouped = df.groupby("Asset Class")["Market Value (CAD)"].sum().reset_index()
grouped.columns = ["Asset Class", "Current $"]
total_value = grouped["Current $"].sum()
grouped["Current %"] = grouped["Current $"] / total_value * 100
grouped = grouped.round(2)
grouped = grouped.set_index("Asset Class").reindex(desired_order).dropna(how='all').reset_index()

# === User selects rebalancing method per class ===
st.subheader("🎯 Allocation Inputs")
st.markdown("<div style='text-align: center; font-size: 1.1rem; font-weight: 500;'>Adjust allocations and toggle locks as needed</div>", unsafe_allow_html=True)
methods, inputs, locks = {}, {}, {}

for asset in desired_order:
    if asset not in grouped["Asset Class"].values:
        continue
    cols = st.columns([1.2, 1.2, 2.2, 1.2], gap="medium")
    with cols[0]:
        st.markdown(f"<div class='block-label' style='text-align:center'>{asset.upper()}</div>", unsafe_allow_html=True)
    with cols[1]:
        methods[asset] = st.selectbox("", ["%", "$", "$ Δ"], key=f"method_{asset}")
    with cols[2]:
        inputs[asset] = st.number_input("", step=100.0 if methods[asset] != "%" else 0.1, key=f"val_{asset}")
    with cols[3]:
        st.markdown("<div class='lock-wrap' style='display: flex; align-items: center; justify-content: center;'>", unsafe_allow_html=True)
        locks[asset] = st.toggle("", value=False, key=f"lock_{asset}")
        st.markdown(f"<span class='lock-icon'>{'🔒' if locks[asset] else '🔓'}</span></div>", unsafe_allow_html=True)

# === Convert all to target $ ===
target_dollars, locked_assets, unlocked_assets = {}, [], []
for _, row in grouped.iterrows():
    asset, current = row["Asset Class"], row["Current $"]
    method, val, locked = methods[asset], inputs[asset], locks[asset]
    if locked:
        if method == "%":
            target_dollars[asset] = val / 100 * total_value
        elif method == "$":
            target_dollars[asset] = val
        elif method == "$ Δ":
            target_dollars[asset] = current + val
        locked_assets.append(asset)
    else:
        unlocked_assets.append(asset)

assigned = sum(target_dollars.values())
remaining_value = total_value - assigned
if unlocked_assets:
    even_value = remaining_value / len(unlocked_assets)
    for asset in unlocked_assets:
        target_dollars[asset] = even_value

# === Final rebalancing plan ===
grouped["Target $"] = grouped["Asset Class"].map(target_dollars)
grouped["Target %"] = grouped["Target $"] / total_value * 100
grouped["Buy/Sell $"] = grouped["Target $"] - grouped["Current $"]
grouped = grouped.round(2)

# === Display table ===
st.subheader("📥 Rebalancing Plan")
display_df = grouped.copy()
for col in ["Current $", "Target $", "Buy/Sell $"]:
    display_df[col] = display_df[col].apply(lambda x: f"${x:,.2f}")
display_df["Current %"] = display_df["Current %"].apply(lambda x: f"{x:.2f}%")
display_df["Target %"] = display_df["Target %"].apply(lambda x: f"{x:.2f}%")
st.dataframe(display_df.set_index("Asset Class"), use_container_width=True)

# === Charts ===
def plot_pie(values, labels, title):
    pairs = [(v, l) for v, l in zip(values, labels) if pd.notnull(v)]
    if not pairs:
        st.warning(f"No data for {title} chart.")
        return
    values, labels = zip(*pairs)
    fig, ax = plt.subplots()
    ax.pie(values, labels=labels, autopct='%1.1f%%', startangle=90, textprops={'color': 'black'})
    ax.axis('equal')
    st.pyplot(fig)

st.subheader("📊 Allocation Charts")
col1, col2 = st.columns(2)
with col1:
    st.markdown("**Current**")
    plot_pie(grouped["Current %"], grouped["Asset Class"], "Current")
with col2:
    st.markdown("**Target**")
    plot_pie(grouped["Target %"], grouped["Asset Class"], "Target")

# === Download file ===
output_path = "rebalancing_plan.xlsx"
grouped.to_excel(output_path, index=False)
with open(output_path, "rb") as f:
    st.download_button("⬇️ Download Excel", f, file_name="rebalancing_plan.xlsx")

# === Summary ===
st.subheader("📝 Summary")
buys = grouped[grouped["Buy/Sell $"] > 0]
sells = grouped[grouped["Buy/Sell $"] < 0]
if buys.empty and sells.empty:
    st.success("✅ No rebalancing needed.")
else:
    col1, col2 = st.columns(2)
    with col1:
        if not buys.empty:
            st.markdown("### 🟢 Buy")
            for _, row in buys.iterrows():
                st.markdown(f"<span style='color:green;'>• **{row['Asset Class']}**: ${row['Buy/Sell $']:,.2f}</span>", unsafe_allow_html=True)
    with col2:
        if not sells.empty:
            st.markdown("### 🔴 Sell")
            for _, row in sells.iterrows():
                st.markdown(f"<span style='color:red;'>• **{row['Asset Class']}**: ${abs(row['Buy/Sell $']):,.2f}</span>", unsafe_allow_html=True)
