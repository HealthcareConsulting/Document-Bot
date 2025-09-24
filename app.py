# app.py — NDIS Doc Bot (Single Client, Hybrid Engine) with Fixed Logo Size Controls

# 1) Imports
import streamlit as st
import importlib.util, sys, json, zipfile
from pathlib import Path
from datetime import datetime

# 2) Load the HYBRID engine (safe text + perfect logo + cover/textbox rescue)
spec = importlib.util.spec_from_file_location(
    "ndis_cli",
    str(Path(__file__).parent / "finalHC.py"),
)
ndis_cli = importlib.util.module_from_spec(spec)
sys.modules["ndis_cli"] = ndis_cli
spec.loader.exec_module(ndis_cli)

# 3) Page setup
st.set_page_config(page_title="NDIS Doc Bot — Single Client", page_icon="🗂️", layout="wide")
st.title("🗂️ NDIS Doc Bot — Single Client (Hybrid)")

# 4) Sidebar: locations & options
with st.sidebar:
    st.header("Locations")
    default_master = (Path.cwd() / "MASTER DOCUMENTS").resolve()
    default_out = (Path.cwd() / "OUTPUT").resolve()

    master_path = st.text_input("Master templates (folder or .zip)", value=str(default_master))
    out_root = st.text_input("Output folder", value=str(default_out))
    
    st.header("Logo Settings")
    st.caption("💡 Control logo sizes for different document areas")
    
    # Initialize session state for logo width if not exists
    if 'logo_width' not in st.session_state:
        st.session_state.logo_width = 25.0
    
    # Display current size prominently
    st.metric("Current Logo Size", f"{st.session_state.logo_width}mm")
    
    # Quick size presets with forced refresh
    st.caption("**Quick presets:**")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Small\n(15mm)", help="Best for headers", key="small_btn"):
            st.session_state.logo_width = 15.0
            st.rerun()
    with col2:
        if st.button("Medium\n(30mm)", help="Good balance", key="medium_btn"):
            st.session_state.logo_width = 30.0
            st.rerun()
    with col3:
        if st.button("Large\n(45mm)", help="Cover pages", key="large_btn"):
            st.session_state.logo_width = 45.0
            st.rerun()
    
    # Manual size input with immediate update
    new_logo_width = st.number_input(
        "Logo width (mm)", 
        min_value=5.0, 
        max_value=80.0, 
        value=float(st.session_state.logo_width),
        step=1.0,
        help="Size applies to both header and cover logos",
        key="logo_width_input"
    )
    
    # Update session state when slider changes
    if abs(new_logo_width - st.session_state.logo_width) > 0.1:
        st.session_state.logo_width = float(new_logo_width)
        st.rerun()
    
    # Size feedback
    if st.session_state.logo_width <= 20:
        size_category = "🟢 Small (good for headers)"
    elif st.session_state.logo_width <= 35:
        size_category = "🔵 Medium (balanced)"
    else:
        size_category = "🟠 Large (good for covers)"
    
    st.info(f"**{st.session_state.logo_width}mm** - {size_category}")
    
    st.header("Options")
    dry_run = st.checkbox("Dry run (no docs written)", value=False)
    st.caption("Tip: Keep defaults if the app lives next to your templates.")

# 5) Helpers
def discover_services(master: Path):
    p = Path(master)
    if p.is_dir():
        return sorted([c.name for c in p.iterdir() if c.is_dir()])
    if p.suffix.lower() == ".zip" and p.exists():
        with zipfile.ZipFile(p, "r") as zf:
            top = {name.split("/")[0] for name in zf.namelist() if "/" in name}
            return sorted(list(top))
    return []

def build_data_dict(basics: dict, extras_text: str):
    data = {}
    for k, v in basics.items():
        if v:
            data[k] = v
    # extras lines as "<key>=value" (angle brackets optional)
    for line in (extras_text or "").splitlines():
        if "=" in line:
            key, val = line.split("=", 1)
            k = key.strip()
            v = val.strip()
            if k and v:
                if not (k.startswith("<") and k.endswith(">")):
                    k = f"<{k.strip('<>')}>"
                data[k] = v
    return data
def title_case_input(label, key):
    """Streamlit text input that always displays and stores title case."""

    def _to_title():
        if st.session_state[key]:
            st.session_state[key] = st.session_state[key].title()

    return st.text_input(label, key=key, on_change=_to_title)



# 6) Client details
st.subheader("Client details")
col1, col2 = st.columns(2)
with col1:
    company_name    = title_case_input("<company name>", key="company_name")
    trading_name    = title_case_input("<trading name>",key="trading name")
    entity_name     = title_case_input("<entity name>",key="entity name")   
    abn             = st.text_input("<abn>")
    acn             = st.text_input("<acn>")          
with col2:
    company_email   = st.text_input("<company email>")
    company_phone   = st.text_input("<company phone>")
    company_address = st.text_input("<company address>")
    website         = st.text_input("<website>")      
    ho              = st.text_input("<ho>")           

st.markdown("**Additional placeholders (optional)** — one per line like `<key>=value`")
extras_text = st.text_area("Extras", height=140, placeholder="<director name>=Jane Doe\n<year>=2025")

# 7) Logo upload
logo_file = st.file_uploader("Upload logo (.png/.jpg)", type=["png","jpg","jpeg"])

# 8) Services - dynamic from local master path
master_path = st.text_input(
    "Enter master templates folder path",
    value=str(Path.cwd() / "MASTER DOCUMENTS")  # or default path you want
)

services_options = discover_services(master_path)
services = st.multiselect(
    "Select services (folders)", 
    options=services_options, 
    default=services_options
)



# 9) Output naming
client_label = st.text_input("Output subfolder name", value=f"CLIENT-{datetime.now().strftime('%Y-%m-%d')}")

st.divider()

# Show current settings before processing
st.info(f"📊 **Current Settings:** Logo size: {st.session_state.logo_width}mm | Dry run: {dry_run}")

colA, colB = st.columns([1,1])
go = colA.button("Generate filled documents" if not dry_run else "Preview (Dry run)", type="primary")
reset = colB.button("Reset form")

if reset:
    # Reset session state
    st.session_state.logo_width = 25.0
    st.rerun()

# 10) Run pipeline
if go:
    master = Path(master_path)
    out_root_p = Path(out_root)
    out_client_dir = out_root_p / client_label

    if not master.exists():
        st.error("Master path does not exist. Fix the path in the sidebar.")
        st.stop()

    out_root_p.mkdir(parents=True, exist_ok=True)
    workdir = out_root_p / "_ui_work"
    if workdir.exists():
        import shutil; shutil.rmtree(workdir)
    workdir.mkdir(parents=True, exist_ok=True)

    # Build data.json
    basics = {
        "<company name>":    company_name,
        "<trading name>":    trading_name,
        "<entity name>":     entity_name,
        "<abn>":             abn,
        "<acn>":             acn,
        "<company email>":   company_email,
        "<company phone>":   company_phone,
        "<company address>": company_address,
        "<website>":         website,
        "<ho>":              ho,
    }
    data_dict = build_data_dict(basics, extras_text)
    tmp_data_json = workdir / "data.json"
    tmp_data_json.write_text(json.dumps(data_dict, indent=2), encoding="utf-8")

    # Save logo
    tmp_logo_path = None
    if logo_file is not None:
        suffix = Path(logo_file.name).suffix.lower() or ".png"
        tmp_logo_path = workdir / f"logo{suffix}"
        tmp_logo_path.write_bytes(logo_file.read())

    services_csv = ",".join(services) if services else ""

    # Display processing info with explicit value
    current_size = float(st.session_state.logo_width)
    st.info(f"📄 Processing with logo size: **{current_size}mm**")
    st.write(f"🔍 Debug: Session state value = {st.session_state.logo_width}")
    st.write(f"🔍 Debug: Converting to float = {current_size}")
    
    with st.spinner("Processing…"):
        # Explicitly pass the current size
        report_path = ndis_cli.run_pipeline(
            master_src=master,
            out_dir=out_client_dir,
            data_json=tmp_data_json,
            logo=tmp_logo_path,
            services_csv=services_csv,
            dry_run=dry_run,
            logo_width_mm=current_size,  # Use the explicit current_size variable
        )

    st.success(f"Done! Logo size used: {current_size}mm")

    # CSV download
    if report_path.exists():
        with open(report_path, "rb") as f:
            st.download_button("Download report CSV", data=f, file_name=report_path.name, mime="text/csv")

    # Zip download (live runs)
    if not dry_run and out_client_dir.exists():
        out_zip = out_client_dir.with_suffix(".zip")
        if out_zip.exists():
            out_zip.unlink()
        import shutil
        shutil.make_archive(str(out_client_dir), "zip", root_dir=out_client_dir)
        with open(out_zip, "rb") as fz:
            st.download_button("Download filled documents (.zip)", data=fz, file_name=out_zip.name, mime="application/zip")

    # Show data.json used
    st.subheader("data.json used")
    st.code(tmp_data_json.read_text(encoding="utf-8"), language="json")

# Status message
st.caption(f"✅ Safe mode active | Current logo size: {st.session_state.logo_width}mm | Header logos capped at 20mm for layout safety")
