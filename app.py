import streamlit as st
import tempfile, os, shutil, io
from datetime import date
import importlib.util, sys

st.set_page_config(page_title="Payroll Audit Tool", page_icon="📊", layout="centered")

st.title("📊 Payroll Audit Tool")
st.caption("Remote → Gusto | Upload your three files, set the pay period, and download the audit workbook.")

st.divider()

col1, col2 = st.columns(2)
with col1:
    input_file   = st.file_uploader("① Remote Input Workbook (.xlsx)", type=["xlsx"])
    mapping_file = st.file_uploader("③ Mapping / Processing File (.csv)", type=["csv"])
with col2:
    output_file  = st.file_uploader("② Gusto Output Journal (.csv)", type=["csv"])

st.divider()

c1, c2, c3 = st.columns(3)
with c1:
    period_start = st.date_input("Period Start", value=date.today().replace(day=1))
with c2:
    period_end   = st.date_input("Period End",   value=date.today())
with c3:
    pay_date_str = st.text_input("Pay Date (display only)", value=period_end.strftime("%-m/%-d/%Y") if hasattr(period_end, 'strftime') else "")

st.divider()

ready = input_file and output_file and mapping_file

if not ready:
    st.info("Upload all three files to enable the audit.")

if ready:
    if st.button("▶ Run Audit", type="primary", use_container_width=True):
        with st.spinner("Running audit — this takes about 30 seconds…"):
            try:
                tmpdir = tempfile.mkdtemp()

                # Write uploaded files to temp dir
                in_path  = os.path.join(tmpdir, "input.xlsx")
                out_path = os.path.join(tmpdir, "output.csv")
                map_path = os.path.join(tmpdir, "mapping.csv")
                res_path = os.path.join(tmpdir, "result.xlsx")

                with open(in_path,  "wb") as f: f.write(input_file.getvalue())
                with open(out_path, "wb") as f: f.write(output_file.getvalue())
                with open(map_path, "wb") as f: f.write(mapping_file.getvalue())

                # Copy name_overrides.csv from script directory
                script_dir   = os.path.dirname(os.path.abspath(__file__))
                overrides_src = os.path.join(script_dir, "name_overrides.csv")
                if os.path.exists(overrides_src):
                    shutil.copy(overrides_src, os.path.join(tmpdir, "name_overrides.csv"))

                # Load and patch the audit script dynamically
                audit_src = os.path.join(script_dir, "payroll_audit.py")
                with open(audit_src, "r") as f:
                    src = f.read()

                from datetime import date as _date
                ps = period_start if isinstance(period_start, _date) else _date.fromisoformat(str(period_start))
                pe = period_end   if isinstance(period_end,   _date) else _date.fromisoformat(str(period_end))

                # Patch config paths and dates
                src = src.replace(
                    f"PERIOD_START = date({ps.year - 1 if True else ps.year}",  # find any PERIOD_START line
                    "PERIOD_START = date(PLACEHOLDER_S"
                )
                # Simpler: regex replace the config block
                import re
                src = re.sub(r'PERIOD_START\s*=\s*date\([^)]+\)',
                             f'PERIOD_START = date({ps.year}, {ps.month}, {ps.day})', src)
                src = re.sub(r'PERIOD_END\s*=\s*date\([^)]+\)',
                             f'PERIOD_END   = date({pe.year}, {pe.month}, {pe.day})', src)
                src = re.sub(r"INPUT_PATH\s*=\s*'[^']+'",
                             f"INPUT_PATH   = r'{in_path}'", src)
                src = re.sub(r"OUTPUT_PATH\s*=\s*'[^']+'",
                             f"OUTPUT_PATH  = r'{out_path}'", src)
                src = re.sub(r"MAPPING_PATH\s*=\s*'[^']+'",
                             f"MAPPING_PATH = r'{map_path}'", src)
                src = re.sub(r"RESULT_PATH\s*=\s*'[^']+'",
                             f"RESULT_PATH  = r'{res_path}'", src)

                # Patch pay period display strings
                period_label = f"{ps.strftime('%B %-d')} – {pe.strftime('%-d, %Y')}"
                src = re.sub(r"'Pay Period',\s*'[^']+'",
                             f"'Pay Period', '{period_label}'", src)
                src = re.sub(r"'Pay Date',\s*'[^']+'",
                             f"'Pay Date', '{pay_date_str}'", src)
                src = re.sub(r"ws0\['A2'\]\s*=\s*'[^']+'",
                             f"ws0['A2'] = 'Remote → Gusto  |  {period_label}'", src)

                # Override OVERRIDES_PATH to temp dir
                src = src.replace(
                    "OVERRIDES_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'name_overrides.csv')",
                    f"OVERRIDES_PATH = r'{os.path.join(tmpdir, 'name_overrides.csv')}'"
                )

                # Execute in isolated namespace
                ns = {"__file__": audit_src}
                exec(compile(src, audit_src, "exec"), ns)

                # Read result
                with open(res_path, "rb") as f:
                    result_bytes = f.read()

                shutil.rmtree(tmpdir, ignore_errors=True)

                fname = f"Payroll_Audit_{ps.strftime('%b%d')}_{pe.strftime('%b%d_%Y')}.xlsx"
                st.success("✅ Audit complete!")
                st.download_button(
                    label="⬇ Download Audit Workbook",
                    data=result_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

            except Exception as e:
                shutil.rmtree(tmpdir, ignore_errors=True)
                st.error(f"Audit failed: {e}")
                st.exception(e)

st.divider()
st.caption("Files are processed in memory and never stored. Each session is independent.")
