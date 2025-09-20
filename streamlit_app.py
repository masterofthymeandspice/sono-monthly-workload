import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import numpy as np

# --------------- App Config ---------------
st.set_page_config(
    page_title="Sono HOH Monthly Workload",
    page_icon="🩻",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --------------- Default Mapping ---------------
DEFAULT_SCAN_TYPE_MAP = {
    'Biopsy Gland, Muscle, Tissue': 'NONE',
    'Consult Radiologist': 'NONE',
    'Ultrasound AAA/IVC Doppler': 'NONE',
    'Ultrasound Abdo/Pelvic/Renal Doppler': 'NONE',
    'Ultrasound LEFT DVT Leg': 'DVT',
    'Ultrasound Neck +/or Thyroid': 'Thyroid',
    'Ultrasound Obstetric Dating': 'First trimester',
    'Ultrasound Pelvis': 'Pelvic and lower GIT',
    'Ultrasound Pelvis NR': 'NONE',
    'Ultrasound Renal Tract': 'Urinary tract',
    'Ultrasound Renal Tract NR': 'Urinary tract',
    'Ultrasound Testes': 'Scrotum',
    'Ultrasound Upper Abdomen': 'Upper abdomen',
    'Ultrasound ^ Chest +/or Abdominal Wall': 'Upper abdomen',
    'Ultrasound ^ Hip + Intervention': 'Drainage/injection/biopsy etc',
    'Ultrasound ^ Intervention Only Hip': 'Drainage/injection/biopsy etc',
    'Ultrasound ^ Intervention Only Knee': 'Drainage/injection/biopsy etc',
    'Ultrasound ^ Intervention Only Shoulder': 'Drainage/injection/biopsy etc',
    'Ultrasound ^ Knee + Intervention': 'Drainage/injection/biopsy etc',
    'Ultrasound ^ Lump or Mass': 'Other',
    'Ultrasound ^ Shoulder + Intervention': 'Drainage/injection/biopsy etc',
    'Ultrasound ^ Unilat Hip +/- Groin': 'MSK other',
    'Ultrasound ^ Unilat Knee': 'MSK other',
    'Ultrasound ^ Unilat Knee NR': 'MSK other',
    'Ultrasound ^ Unilat Shoulder/Upper Arm': 'Shoulder',
    'Ultrasound to Guide Cannulation': 'Drainage/injection/biopsy etc',
    'Ultrasound to Guide Surgical Procedure': 'Drainage/injection/biopsy etc',
    'Venous Reflux or Obstruction': 'NONE'
}

GROUP_OPTIONS = sorted(list({
    'Pelvic and lower GIT',
    'Drainage/injection/biopsy etc',
    'Upper abdomen',
    'Shoulder',
    'Urinary tract',
    'Thyroid',
    'Other',
    'MSK other',
    'First trimester',
    'Scrotum',
    'DVT',
    'Unassigned',
    'NONE'
}))

# --------------- Helpers ---------------
def dict_to_df(mapping: dict) -> pd.DataFrame:
    if not mapping:
        return pd.DataFrame(columns=["Exam Description", "Scan Type Group"])
    return pd.DataFrame(
        [{"Exam Description": k, "Scan Type Group": v} for k, v in mapping.items()]
    ).sort_values("Exam Description").reset_index(drop=True)

def df_to_dict(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {}
    clean = df.dropna(subset=["Exam Description"])
    result = {}
    for _, row in clean.iterrows():
        result[str(row["Exam Description"]).strip()] = None if pd.isna(row["Scan Type Group"]) else str(row["Scan Type Group"]).strip()
    return result

def process_data(
    df: pd.DataFrame,
    scan_type_map: dict,
    first_week_date: pd.Timestamp,
    hours_per_record: float,
    visit_number_col: str = "Visit Number",
    visit_start_col: str = "Visit Start Date And Time",
    exam_description_col: str = "Exam Description",
):
    # Forward-fill merged cells
    for col in [visit_number_col, visit_start_col]:
        if col in df.columns:
            df[col] = df[col].ffill()

    # Ensure datetime parsing
    if visit_start_col in df.columns:
        # Try known format first; coerce if not matching
        df[visit_start_col] = pd.to_datetime(
            df[visit_start_col],
            format="%d/%m/%Y %H:%M:%S",
            errors="coerce"
        )
        # If many NaT (format mismatch), try generic parsing
        if df[visit_start_col].isna().mean() > 0.5:
            df[visit_start_col] = pd.to_datetime(df[visit_start_col], errors="coerce")

    # Map Scan Type Group
    if exam_description_col not in df.columns:
        raise KeyError(f"Expected column '{exam_description_col}' not found in the uploaded file.")
    df["Scan Type Group"] = df[exam_description_col].map(scan_type_map)

    # Week Number
    def compute_week(dt):
        if pd.notnull(dt) and dt >= first_week_date:
            return ( (dt - first_week_date).days // 7 ) + 1
        return None

    df["Week Number"] = df[visit_start_col].apply(compute_week) if visit_start_col in df.columns else None

    # Hours
    df["Hours"] = float(hours_per_record)

    # Filter and de-duplicate
    filtered_df = df[df["Scan Type Group"] != "NONE"].copy()
    if visit_number_col not in df.columns:
        raise KeyError(f"Expected column '{visit_number_col}' not found in the uploaded file.")
    deduped_df = filtered_df.drop_duplicates(subset=visit_number_col, keep="first").copy()

    # Pivot
    pivot_df = pd.pivot_table(
        deduped_df,
        index="Scan Type Group",
        columns="Week Number",
        values="Hours",
        aggfunc="sum",
        fill_value=0,
    )
    # Keep columns sorted numerically if present
    try:
        pivot_df = pivot_df.reindex(sorted(pivot_df.columns.dropna()), axis=1)
    except Exception:
        pass

    return df, filtered_df, deduped_df, pivot_df

def build_excel_bytes(df_original, df_deduped, pivot_df, file_basename: str = "summary", chart_png_bytes: bytes = None) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_original.to_excel(writer, sheet_name="original", index=False)
        df_deduped.to_excel(writer, sheet_name="Deduped Data", index=False)
        pivot_df.to_excel(writer, sheet_name="Pivot Summary")
        if chart_png_bytes:
            wb = writer.book
            ws = wb.create_sheet(title="Charts")
            pil_img = PILImage.open(BytesIO(chart_png_bytes))
            img = XLImage(pil_img)
            ws.add_image(img, "A1")
    output.seek(0)
    return output.getvalue()

# --------------- Sidebar (Controls) ---------------
with st.sidebar:
    st.header("⚙️ Options")

    st.caption("Excel read options")
    header_row = st.number_input("Header row (1-indexed)", min_value=1, value=17, step=1, help="Row number that contains column headers.")
    sheet_index = st.number_input("Sheet index (0-based)", min_value=0, value=0, step=1)

    st.divider()
    st.caption("Processing options")
    hours_per_record = st.number_input("Hours per record", min_value=0.0, value=0.75, step=0.25)
    first_week_date = st.date_input("First week of June 2025", value=pd.Timestamp("2025-06-01").date())
    first_week_ts = pd.Timestamp(first_week_date)

    st.divider()
    st.caption("Column names (ensure they match your file)")
    visit_number_col = st.text_input("Visit Number column", value="Visit Number")
    visit_start_col = st.text_input("Visit Start Date And Time column", value="Visit Start Date And Time")
    exam_desc_col = st.text_input("Exam Description column", value="Exam Description")

# --------------- Header ---------------
st.markdown(
    """
    <style>
    .badge {background: #eef2ff; color:#3730a3; padding:2px 8px; border-radius:999px; font-size:0.8rem; margin-left:8px;}
    .step {border-left: 4px solid #22c55e; padding-left:12px; margin: 10px 0;}
    .muted {color:#64748b;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🩻 Sono HOH Monthly Workload")
st.markdown("<span class='muted'>Upload → Edit Mapping → Process → Preview → Download</span>", unsafe_allow_html=True)

# --------------- File Upload ---------------
uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xlsm"],
    accept_multiple_files=False,
    help="The app reads the first sheet by default (configurable in the sidebar)."
)

# --------------- Mapping (initialize and augment from uploaded file) ---------------
if "scan_map_df" not in st.session_state:
    st.session_state.scan_map_df = dict_to_df(DEFAULT_SCAN_TYPE_MAP)

# If a file is uploaded, scan unique Exam Descriptions and add missing rows with default 'Other'
if uploaded_file is not None:
    try:
        preview_df = pd.read_excel(
            uploaded_file,
            sheet_name=int(sheet_index),
            header=int(header_row) - 1
        )
        uploaded_file.seek(0)
        col_name = exam_desc_col.strip()
        if col_name in preview_df.columns:
            unique_desc = (
                preview_df[col_name]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
            )
            current_map_df = st.session_state.scan_map_df.copy()
            if not current_map_df.empty and "Exam Description" in current_map_df.columns:
                existing = set(current_map_df["Exam Description"].astype(str).str.strip())
            else:
                existing = set()
            new_items = [d for d in unique_desc if d not in existing]
            if new_items:
                add_df = pd.DataFrame(
                    [{"Exam Description": d, "Scan Type Group": "Unassigned"} for d in new_items]
                )
                st.session_state.scan_map_df = (
                    pd.concat([current_map_df, add_df], ignore_index=True)
                    .sort_values("Exam Description")
                    .reset_index(drop=True)
                )
                st.info(f"Detected {len(new_items)} new Exam Description value(s) in the uploaded file. Added to the mapping as 'Unassigned'.")
                st.caption("Newly detected Exam Descriptions: " + ", ".join(sorted(map(str, new_items))))
        else:
            st.warning(f"Column '{col_name}' not found in uploaded file; cannot suggest mapping rows.")
    except Exception as e:
        st.warning(f"Unable to inspect uploaded file for mapping suggestions: {e}")

# --------------- Mapping Editor ---------------
with st.expander("✏️ Edit scan_type_map (Exam Description → Scan Type Group)", expanded=True):
    st.caption("Tip: You can add/remove rows. 'NONE' will be filtered out before deduping.")
    edited_map_df = st.data_editor(
        st.session_state.scan_map_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Exam Description": st.column_config.TextColumn(
                "Exam Description",
                help="Exact text as it appears in the Excel file."
            ),
            "Scan Type Group": st.column_config.TextColumn(
                "Scan Type Group",
                help="Enter any group name. Use 'NONE' to exclude from analysis; 'Unassigned' if undecided."
            ),
        },
        key="scan_map_editor"
    )
    # Keep the edited map in session
    st.session_state.scan_map_df = edited_map_df

# --------------- Process Button ---------------
process_btn = st.button(
    "🚀 Process Data",
    use_container_width=True,
    type="primary",
    disabled=uploaded_file is None
)

# --------------- Processing ---------------
if process_btn:
    if uploaded_file is None:
        st.warning("Please upload a file first.")
        st.stop()

    with st.spinner("Reading and processing..."):
        try:
            # Convert mapping edits to dict
            scan_type_map = df_to_dict(st.session_state.scan_map_df)

            # Read Excel (header is 0-based in pandas; user input is 1-based)
            uploaded_file.seek(0)
            df = pd.read_excel(
                uploaded_file,
                sheet_name=int(sheet_index),
                header=int(header_row) - 1
            )

            original_df, filtered_df, deduped_df, pivot_df = process_data(
                df=df,
                scan_type_map=scan_type_map,
                first_week_date=first_week_ts,
                hours_per_record=hours_per_record,
                visit_number_col=visit_number_col.strip(),
                visit_start_col=visit_start_col.strip(),
                exam_description_col=exam_desc_col.strip(),
            )

            # Stats
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Rows (original)", f"{len(original_df):,}")
            c2.metric("Rows after 'NONE' filter", f"{len(filtered_df):,}")
            c3.metric("Rows after de-dup", f"{len(deduped_df):,}")
            total_hours = deduped_df["Hours"].sum() if "Hours" in deduped_df.columns else 0
            c4.metric("Total hours (deduped)", f"{total_hours:,.2f}")

            # Prepare data for charts
            chart_df = deduped_df.dropna(subset=["Week Number", "Scan Type Group"]).copy()
            chart_png_bytes = None
            if not chart_df.empty:
                agg = chart_df.groupby(["Week Number", "Scan Type Group"], as_index=False)["Hours"].sum()
                week_totals_series = agg.groupby("Week Number")["Hours"].sum()
                week_totals = week_totals_series.to_dict()
                agg["Week Total"] = agg["Week Number"].map(week_totals)
                agg["Percent"] = (agg["Hours"] / agg["Week Total"] * 100).fillna(0.0)
                week_order = sorted(week_totals.keys())
                week_label_map = {wk: f"Week {int(wk)} ({week_totals[wk]:.1f}hrs)" for wk in week_order}
                agg["Week Label"] = agg["Week Number"].map(week_label_map)
                week_label_order = [week_label_map[wk] for wk in week_order]

                # Plotly stacked proportion chart
                fig_plotly = px.bar(
                    agg,
                    x="Week Label",
                    y="Hours",
                    color="Scan Type Group",
                    category_orders={"Week Label": week_label_order},

                    custom_data=["Scan Type Group", "Week Number", "Week Total", "Hours", "Percent"],
                )
                fig_plotly.update_layout(
                    barmode="stack",
                    barnorm="percent",
                    xaxis_title="Week",
                    yaxis_title="Proportion (%)",
                    legend_title="Scan Type Group",
                    margin=dict(l=10, r=10, t=30, b=10),
                    hovermode="x",
                )
                fig_plotly.update_traces(
                    hovertemplate="<b>%{customdata[0]}</b><br>Week %{customdata[1]} (%{customdata[2]:.1f}hrs)<br>Hours: %{customdata[3]:.2f}hrs<br>Proportion: %{customdata[4]:.1f}%<extra></extra>"
                )

                # Matplotlib stacked proportion chart (high-res)
                pivot_hours = agg.pivot(index="Week Label", columns="Scan Type Group", values="Hours").fillna(0.0)
                prop = pivot_hours.div(pivot_hours.sum(axis=1), axis=0).fillna(0.0) * 100.0
                prop = prop.reindex(index=week_label_order)
                fig, ax = plt.subplots(figsize=(12, 6), dpi=200)
                bottom = np.zeros(len(prop))
                for group in prop.columns:
                    vals = prop[group].values
                    ax.bar(prop.index, vals, bottom=bottom, label=group)
                    bottom += vals
                ax.set_ylabel("Proportion (%)")
                ax.set_xlabel("Week")
                ax.set_title("Scan Type Group by Week (Proportion)")
                ax.set_ylim(0, 100)
                ax.tick_params(axis="x", rotation=30)
                ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0), title="Scan Type Group")
                fig.tight_layout()

                # Save matplotlib chart to PNG bytes
                _buf = BytesIO()
                fig.savefig(_buf, format="png", dpi=200, bbox_inches="tight")
                _buf.seek(0)
                chart_png_bytes = _buf.getvalue()

            # Tabs for previews and charts
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["📄 Original", "🧹 Deduped", "📊 Pivot Summary", "📈 Plotly Proportions", "🖼️ Matplotlib Proportions"])
            with tab1:
                st.dataframe(original_df, use_container_width=True, hide_index=True)
            with tab2:
                st.dataframe(deduped_df, use_container_width=True, hide_index=True)
            with tab3:
                st.dataframe(pivot_df, use_container_width=True)
            with tab4:
                if chart_df.empty:
                    st.info("No chart data available. Ensure 'Week Number' and 'Scan Type Group' are present.")
                else:
                    st.plotly_chart(fig_plotly, use_container_width=True)
            with tab5:
                if chart_df.empty:
                    st.info("No chart data available to render.")
                else:
                    st.pyplot(fig, use_container_width=True)

            # Build Excel for download
            base_name = uploaded_file.name.rsplit(".", 1)[0]
            download_name = f"{base_name}_summary.xlsx"
            excel_bytes = build_excel_bytes(original_df, deduped_df, pivot_df, file_basename=base_name, chart_png_bytes=chart_png_bytes)

            st.success("Processing complete. Download your results below.")
            st.download_button(
                label="⬇️ Download Excel Summary",
                data=excel_bytes,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.balloons()

        except KeyError as e:
            st.error(f"Missing expected column: {e}")
        except Exception as e:
            st.exception(e)

# --------------- Footer ---------------
st.markdown("---")
st.caption("Built with Streamlit and pandas | Mapping is editable before processing | Hours fixed per record with a configurable value")
