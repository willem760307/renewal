import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import uuid
import extra_streamlit_components as stx

# --- Configuration & Secrets ---
st.set_page_config(page_title="續約管理", layout="wide")

# Google Sheets Initialization
def get_gsheet_client():
    try:
        # Expected st.secrets["gcp_service_account"] to be a dict
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"無法連線至 Google Sheets: {e}")
        return None

def get_worksheets():
    client = get_gsheet_client()
    if not client:
        return None, None, None
    try:
        sh = client.open_by_key(st.secrets["spreadsheet_id"])
        main_ws = sh.worksheet("MainData")
        comment_ws = sh.worksheet("Comments")
        
        # Create Checklist worksheet if not exists
        try:
            checklist_ws = sh.worksheet("Checklist")
        except gspread.WorksheetNotFound:
            checklist_ws = sh.add_worksheet(title="Checklist", rows="1000", cols="5")
            checklist_ws.append_row(["關聯物件地址", "房東資料", "房客資料", "物件資料", "安全檢核"])
            
        return main_ws, comment_ws, checklist_ws
    except Exception as e:
        st.error(f"找不到工作表: {e}")
        return None, None, None

# --- Helpers ---
def load_main_data(main_ws):
    data = main_ws.get_all_records()
    df = pd.DataFrame(data)
    
    # Fix phone numbers if they lost their leading 0 upon loading from Google Sheets
    phone_keywords = ["手機", "電話", "phone", "聯絡"]
    for col in df.columns:
        if any(k in str(col).lower() for k in phone_keywords):
            def fix_phone(val):
                if pd.isna(val) or val == "": return ""
                val_str = str(val).split('.')[0]
                digits = ''.join(filter(str.isdigit, val_str))
                if len(digits) == 9 and digits.startswith('9'):
                    return '0' + digits
                return val_str
            df[col] = df[col].apply(fix_phone)
            
    return df

def load_comments(comment_ws):
    try:
        data = comment_ws.get_all_records()
    except Exception:
        data = []
    
    df = pd.DataFrame(data)
    expected_cols = ["關聯物件地址", "留言時間", "留言內容", "留言ID"]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = pd.Series(dtype=str)
    return df

def load_checklist(checklist_ws):
    try:
        data = checklist_ws.get_all_records()
    except Exception:
        data = []
    
    df = pd.DataFrame(data)
    expected_cols = ["關聯物件地址", "房東資料", "房客資料", "物件資料", "安全檢核", "證件期限檢核", "狀態", "已報業績"]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = pd.Series(dtype=str)
    return df

def save_to_main(main_ws, df):
    # Fix for InvalidJSONError: NaN values are not JSON compliant
    # Also convert everything to string to ensure gspread can handle it
    df_clean = df.fillna("").astype(str)
    main_ws.clear()
    main_ws.update([df_clean.columns.values.tolist()] + df_clean.values.tolist())

def get_status_emoji(status):
    mapping = {
        "未送預審": "⚪",
        "預審中": "🟡",
        "預審通過": "🔵",
        "簽約中": "🟣",
        "已完成": "🟢",
        "不續約": "🔴",
    }
    return mapping.get(status, "⚪")

def render_inline_comments(item_address, comment_ws, is_dialog=False):
    if "df_comments" not in st.session_state: return
    full_c = st.session_state.df_comments.copy()
    address_comments = full_c[full_c['關聯物件地址'] == item_address].copy()
    prefix = "dlg_" if is_dialog else "tbl_"
    
    if not address_comments.empty:
        address_comments['留言時間'] = pd.to_datetime(address_comments['留言時間'], errors='coerce')
        address_comments = address_comments.sort_values(by='留言時間', ascending=True)
        for _, c_row in address_comments.iterrows():
            cid = c_row['留言ID']
            edit_key = f"{prefix}editing_{cid}"
            
            with st.container(border=True):
                if st.session_state.get(edit_key, False):
                    st.caption(f"🕒 {c_row['留言時間']}")
                    new_text = st.text_area("編輯備註", value=c_row['留言內容'], key=f"{prefix}edit_input_{cid}", label_visibility="collapsed")
                    c1, c2 = st.columns([1, 10])
                    if c1.button("儲存修改", key=f"{prefix}save_{cid}"):
                        idx = full_c[full_c['留言ID'] == cid].index[0]
                        full_c.at[idx, '留言內容'] = new_text
                        st.session_state.df_comments = full_c
                        st.session_state[edit_key] = False
                        st.rerun()
                    c2.button("取消", key=f"{prefix}cancel_{cid}", on_click=lambda k=edit_key: st.session_state.update({k: False}))
                else:
                    c_txt, c_btn1, c_btn2 = st.columns([8.5, 0.75, 0.75])
                    escaped_content = str(c_row['留言內容']).replace('\n', '<br>')
                    c_txt.markdown(f"{escaped_content} <span style='font-size:0.8em; color:#a0aec0; margin-left:12px;'>🕒 {c_row['留言時間']}</span>", unsafe_allow_html=True)
                    c_btn1.button("📝", key=f"{prefix}edit_btn_{cid}", help="編輯", on_click=lambda k=edit_key: st.session_state.update({k: True}), use_container_width=True)
                    if c_btn2.button("🗑️", key=f"{prefix}del_btn_{cid}", help="刪除", use_container_width=True):
                        idx = full_c[full_c['留言ID'] == cid].index[0]
                        full_c = full_c.drop(idx)
                        st.session_state.df_comments = full_c
                        st.toast("備註已刪除 (需手動儲存同步)")
                        st.rerun()

    def submit_inline_comment(addr, key):
        val = st.session_state.get(key, "").strip()
        if val:
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_id = uuid.uuid4().hex[:8]
            f_c = st.session_state.df_comments
            new_row = [addr, now, val, new_id]
            new_df = pd.DataFrame([new_row], columns=f_c.columns)
            st.session_state.df_comments = pd.concat([f_c, new_df], ignore_index=True)
            st.session_state[key] = ""
            st.toast("留言暫存成功 (記得點擊表格右側或視窗下方的儲存按鈕同步)")

    c_in, c_btn = st.columns([9, 1])
    input_key = f"{prefix}input_new_note_{item_address}"
    c_in.text_input("快速新增備註", key=input_key, label_visibility="collapsed", placeholder="💬 在此新增留言/備註...")
    c_btn.button("送出", key=f"{prefix}btn_send_{item_address}", on_click=submit_inline_comment, args=(item_address, input_key))

# --- Authentication ---
def get_cookie_manager():
    return stx.CookieManager(key="cookie_mgr")

def check_password():
    """Returns True if the user had the correct password or active cookie."""
    cookie_manager = get_cookie_manager()
    
    # Check if they already have an active auth cookie
    if cookie_manager.get(cookie="auth_token") == "authenticated":
        return True

    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            # Set cookie for 365 days
            expiry = datetime.datetime.now() + datetime.timedelta(days=365)
            cookie_manager.set("auth_token", "authenticated", expires_at=expiry)
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("請輸入系統密碼 (登入後此電腦將記憶一年)", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("請輸入系統密碼 (登入後此電腦將記憶一年)", type="password", on_change=password_entered, key="password")
        st.error("😕 密碼錯誤")
        return False
    else:
        return True

@st.dialog("➕ 新增物件", width="large")
def show_add_property_dialog(df, main_ws, address_col, expiry_col, display_fields, other_fields):
    st.write("請輸入新增物件的基本資料，【地址】為必填項目。")
    
    with st.form("add_prop_form"):
        new_addr = st.text_input(f"📍 {address_col} (必填)*")
        new_date = st.date_input(f"📅 {expiry_col} (必填)*")
        
        st.markdown("#### 主要欄位 (選填)")
        cols = st.columns(2)
        new_vals = {}
        for i, field in enumerate(display_fields):
            if field not in [address_col, expiry_col]:
                new_vals[field] = cols[i % 2].text_input(f"{field}")
                
        st.markdown("#### 完整詳細欄位 (選填)")
        with st.expander("展開完整欄位"):
            e_cols = st.columns(2)
            for i, field in enumerate(other_fields):
                if field not in [address_col, expiry_col]:
                    new_vals[field] = e_cols[i % 2].text_input(f"{field}")
                
        if st.form_submit_button("💾 儲存新增物件", type="primary", use_container_width=True):
            if not new_addr.strip():
                st.error(f"【{address_col}】不能為空！請填寫地址。")
            else:
                new_row = {c: "" for c in df.columns}
                new_row[address_col] = new_addr.strip()
                new_row[expiry_col] = new_date.strftime("%Y-%m-%d") if hasattr(new_date, "strftime") else str(new_date)
                
                for k, v in new_vals.items():
                    new_row[k] = v.strip()
                    
                if '關聯物件地址' in df.columns:
                    city = new_row.get('物件縣市', '')
                    addr = new_row.get('租賃地址', new_addr.strip())
                    new_row['關聯物件地址'] = str(city) + str(addr)
                    if not new_row['關聯物件地址'] or new_row['關聯物件地址'] == "None":
                        new_row['關聯物件地址'] = new_addr.strip()
                
                updated_df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_to_main(main_ws, updated_df)
                st.session_state.df_main = updated_df
                st.success("物件已成功新增！")
                st.rerun()

@st.dialog("📍 依縣市分群清單")
def show_grouped_addresses(df, address_col):
    st.write("此處顯示您**目前篩選月份**的地址分群，可直接點擊右上角的複製按鈕。")
    
    addresses = df[address_col].dropna().astype(str).tolist()
    taipei_list = [a for a in addresses if "台北市" in a or "臺北市" in a]
    nt_list = [a for a in addresses if "新北市" in a]
    other_list = [a for a in addresses if "台北市" not in a and "臺北市" not in a and "新北市" not in a]
    
    st.subheader(f"🏙️ 台北市 ({len(taipei_list)}件)")
    if taipei_list:
        st.code("\n".join(taipei_list), language="text")
    else:
        st.write("無")
        
    st.subheader(f"🌇 新北市 ({len(nt_list)}件)")
    if nt_list:
        st.code("\n".join(nt_list), language="text")
    else:
        st.write("無")

    if other_list:
        st.subheader(f"🌍 其他縣市 ({len(other_list)}件)")
        st.code("\n".join(other_list), language="text")

@st.dialog("📋 物件完整明細", width="large")
def show_property_details(row, address_col, display_fields, other_fields, comment_ws, checklist_ws, main_ws):
    item_address = str(row[address_col])
    
    # Custom layout grouper to ensure names and numbers sit side by side natively in the iter loop
    def pair_related_fields(fields):
        ordered = []
        # Preferred left-right pairs visually
        pairs = [("房東姓名", "房東電話"), ("房客姓名", "房客電話"), ("連絡人", "連絡人電話"), ("代理人", "代理人電話")]
        placed = set()
        for p1, p2 in pairs:
            if p1 in fields and p2 in fields:
                ordered.extend([p1, p2])
                placed.add(p1)
                placed.add(p2)
        # Anything leftover appended naturally
        for f in fields:
            if f not in placed:
                ordered.append(f)
        return ordered
        
    display_fields = pair_related_fields(display_fields)
    other_fields = pair_related_fields(other_fields)
    
    def format_tenant_identity(val):
        val_str = str(val).replace('.0', '').strip()
        if val_str == "0" or "一般" in val_str:
            return "一般戶"
        elif val_str == "1" or "第一" in val_str:
            return "第一類弱勢戶"
        elif val_str == "2" or "第二" in val_str:
            return "第二類弱勢戶"
        return val

    idx = row.name
    edit_state_key = f"editing_props_{item_address}"
    is_editing = st.session_state.get(edit_state_key, False)
    
    c_head1, c_head2 = st.columns([8, 2])
    if not is_editing:
        c_head2.button("✏️ 編輯明細", use_container_width=True, on_click=lambda k=edit_state_key: st.session_state.update({k: True}))
            
        cols = st.columns(2)
        for i, field in enumerate(display_fields):
            val = row[field]
            if "身分" in field:
                val = format_tenant_identity(val)
            cols[i % 2].write(f"**{field}**: {val}")
        
        st.divider()
        st.markdown("#### 📖 完整欄位資料")
        e_cols = st.columns(2)
        for i, field in enumerate(other_fields):
            val = row[field]
            if "身分" in field:
                val = format_tenant_identity(val)
            e_cols[i % 2].write(f"**{field}**: {val}")

    else:
        c_head2.button("❌ 取消編輯", use_container_width=True, on_click=lambda k=edit_state_key: st.session_state.update({k: False}))
            
        with st.form(f"edit_form_{item_address}"):
            new_vals = {}
            st.markdown("#### ✏️ 編輯主要欄位")
            cols = st.columns(2)
            for i, field in enumerate(display_fields):
                current_val = str(row[field]) if pd.notna(row[field]) else ""
                new_vals[field] = cols[i % 2].text_input(f"{field}", value=current_val)
                
            st.markdown("#### 📖 編輯完整欄位")
            e_cols = st.columns(2)
            for i, field in enumerate(other_fields):
                current_val = str(row[field]) if pd.notna(row[field]) else ""
                new_vals[field] = e_cols[i % 2].text_input(f"{field}", value=current_val)
                
            if st.form_submit_button("💾 儲存並離開", type="primary", use_container_width=True):
                df_main_ref = st.session_state.df_main.copy()
                for k, v in new_vals.items():
                    df_main_ref.at[idx, k] = v.strip()
                
                # Update relation key if address edited
                if new_vals.get(address_col, item_address).strip() != item_address:
                    city = df_main_ref.at[idx, '物件縣市'] if '物件縣市' in df_main_ref.columns else ""
                    addr = df_main_ref.at[idx, '租賃地址'] if '租賃地址' in df_main_ref.columns else new_vals.get(address_col, "").strip()
                    new_rel_addr = str(city) + str(addr)
                    df_main_ref.at[idx, '關聯物件地址'] = new_rel_addr if new_rel_addr else new_vals.get(address_col, "").strip()
                    
                save_to_main(main_ws, df_main_ref)
                st.session_state.df_main = df_main_ref
                st.session_state[f"editing_props_{item_address}"] = False
                st.success("明細已更新！")
                st.rerun()

    st.divider()
    st.markdown("✅ **確認檢核表**")
    import uuid
    from datetime import datetime as dt
    
    full_chk = st.session_state.df_checklist.copy()
    chk_row_idx = full_chk.index[full_chk['關聯物件地址'] == item_address].tolist()
    if chk_row_idx:
        chk_data = full_chk.iloc[chk_row_idx[0]]
        chk1_val = str(chk_data.get("房東資料", "")) == "True"
        chk2_val = str(chk_data.get("房客資料", "")) == "True"
        chk3_val = str(chk_data.get("物件資料", "")) == "True"
        chk4_val = str(chk_data.get("安全檢核", "")) == "True"
        chk5_val = str(chk_data.get("證件期限檢核", "")) == "True"
        chk_status_val = str(chk_data.get("狀態", "未送預審"))
        chk_perf_val = str(chk_data.get("已報業績", "")) == "True"
    else:
        chk1_val, chk2_val, chk3_val, chk4_val, chk5_val, chk_perf_val = False, False, False, False, False, False
        chk_status_val = "未送預審"

    def dlg_update_checklist(addr, field, key):
        if checklist_ws is None: return
        new_val = st.session_state[key]
        f_chk = st.session_state.df_checklist.copy()
        idx_list = f_chk.index[f_chk['關聯物件地址'] == addr].tolist()
        if not idx_list:
            new_row = {"關聯物件地址": addr, "房東資料": "", "房客資料": "", "物件資料": "", "安全檢核": "", "證件期限檢核": "", "狀態": "未送預審", "已報業績": "False"}
            new_row[field] = str(new_val)
            f_chk = pd.concat([f_chk, pd.DataFrame([new_row])], ignore_index=True)
        else:
            f_chk.at[idx_list[0], field] = str(new_val)
        st.session_state.df_checklist = f_chk

    c_chk1, c_chk2, c_chk3, c_chk4 = st.columns(4)
    c_chk1.checkbox("房東資料", value=chk1_val, key=f"dlg_chk1_{item_address}", on_change=dlg_update_checklist, args=(item_address, "房東資料", f"dlg_chk1_{item_address}"))
    
    tenant_val = ""
    for c in row.index:
        if "身分" in c:
            tenant_val = str(row[c])
            break
    is_class_1 = "1" in tenant_val or "第一" in tenant_val
    is_class_2 = "2" in tenant_val or "第二" in tenant_val
    needs_expiry_check = is_class_1 or is_class_2
    
    c_chk2.checkbox("房客資料", value=chk2_val, key=f"dlg_chk2_{item_address}", on_change=dlg_update_checklist, args=(item_address, "房客資料", f"dlg_chk2_{item_address}"))
    if needs_expiry_check:
        expiry_label = "🚨 更新證件期限" if is_class_2 else "🚨 注意證件期限"
        c_chk2.checkbox(expiry_label, value=chk5_val, key=f"dlg_chk5_{item_address}", on_change=dlg_update_checklist, args=(item_address, "證件期限檢核", f"dlg_chk5_{item_address}"))
        
    c_chk3.checkbox("物件資料", value=chk3_val, key=f"dlg_chk3_{item_address}", on_change=dlg_update_checklist, args=(item_address, "物件資料", f"dlg_chk3_{item_address}"))
    c_chk4.checkbox("安全檢核", value=chk4_val, key=f"dlg_chk4_{item_address}", on_change=dlg_update_checklist, args=(item_address, "安全檢核", f"dlg_chk4_{item_address}"))

    c_s1, c_s2 = st.columns(2)
    status_options = ["未送預審", "預審中", "預審通過", "簽約中", "已完成"]
    current_status = chk_status_val if chk_status_val in status_options else "未送預審"
    c_s1.selectbox("目前狀態", options=status_options, index=status_options.index(current_status), key=f"dlg_status_{item_address}", on_change=dlg_update_checklist, args=(item_address, "狀態", f"dlg_status_{item_address}"))
    c_s2.checkbox("已報業績", value=chk_perf_val, key=f"dlg_perf_{item_address}", on_change=dlg_update_checklist, args=(item_address, "已報業績", f"dlg_perf_{item_address}"))

    st.divider()
    st.markdown("📝 **備註區**")
    
    full_c = load_comments(comment_ws)
    address_comments = full_c[full_c['關聯物件地址'] == item_address].copy()
    
    if not address_comments.empty:
        address_comments['留言時間'] = pd.to_datetime(address_comments['留言時間'], errors='coerce')
        address_comments = address_comments.sort_values(by='留言時間', ascending=True)
        for _, c_row in address_comments.iterrows():
            with st.container(border=True):
                st.caption(f"🕒 {c_row['留言時間']}")
                
                if st.session_state.get(f"editing_{c_row['留言ID']}", False):
                    new_text = st.text_area("編輯備註", value=c_row['留言內容'], key=f"edit_input_{c_row['留言ID']}")
                    c1, c2 = st.columns(2)
                    if c1.button("儲存修改", key=f"save_{c_row['留言ID']}"):
                        idx = full_c[full_c['留言ID'] == c_row['留言ID']].index[0]
                        full_c.at[idx, '留言內容'] = new_text
                        comment_ws.clear()
                        comment_ws.update([full_c.columns.values.tolist()] + full_c.values.tolist())
                        st.session_state[f"editing_{c_row['留言ID']}"] = False
                        st.rerun()
                    if c2.button("取消", key=f"cancel_{c_row['留言ID']}"):
                        st.session_state[f"editing_{c_row['留言ID']}"] = False
                        st.rerun()
                else:
                    st.write(c_row['留言內容'])
                    c1, c2, _ = st.columns([0.1, 0.1, 0.8])
                    if c1.button("📝", key=f"edit_{c_row['留言ID']}", help="編輯"):
                        st.session_state[f"editing_{c_row['留言ID']}"] = True
                        st.rerun()
                    if c2.button("🗑️", key=f"del_{c_row['留言ID']}", help="刪除"):
                        idx = full_c[full_c['留言ID'] == c_row['留言ID']].index[0]
                        full_c = full_c.drop(idx)
                        comment_ws.clear()
                        if not full_c.empty:
                            comment_ws.update([full_c.columns.values.tolist()] + full_c.values.tolist())
                        st.toast("備註已刪除")
                        st.rerun()

    # Add New Comment
    new_comment = st.text_area("新增備註...", key=f"input_new_note_{item_address}")
    if st.button("送出", key=f"btn_send_{item_address}"):
        if new_comment.strip():
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_id = uuid.uuid4().hex[:8]
            new_row = [item_address, now, new_comment.strip(), new_id]
            comment_ws.append_row(new_row)
            st.success("備註已新增！")
            st.rerun()

# --- Main App Logic ---
def main():
    if not check_password():
        st.stop()

    st.title("📂 續約管理")
    
    # --- Modern UI/UX CSS Injection ---
    st.markdown("""
    <style>
        /* Typography */
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600;700&display=swap');
        html, body, [class*="css"]  {
            font-family: 'Noto Sans TC', sans-serif;
        }
        
        /* App Background */
        .stApp {
            background-color: #f7f9fc;
        }
        
        /* Inputs & Selectboxes */
        .stTextInput>div>div>input, .stSelectbox>div>div>select, .stMultiSelect>div>div>div {
            border-radius: 12px !important;
            border: 1px solid #e2e8f0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02) !important;
            transition: all 0.3s ease;
            background-color: white;
        }
        .stTextInput>div>div>input:focus, .stSelectbox>div>div>select:focus {
            border-color: #3b82f6 !important;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15) !important;
        }
        
        /* Buttons */
        button[kind="secondary"] {
            border-radius: 12px !important;
            font-weight: 500 !important;
            background-color: white !important;
            border: 1px solid #e2e8f0 !important;
            transition: all 0.2s ease !important;
            color: #475569 !important;
        }
        button[kind="secondary"]:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 15px rgba(0,0,0,0.06) !important;
            color: #2563eb !important;
            border-color: #aec0d6 !important;
        }
        
        /* Primary Buttons */
        button[kind="primary"] {
            border-radius: 12px !important;
            font-weight: 600 !important;
            transition: all 0.2s ease !important;
            box-shadow: 0 4px 10px rgba(255, 75, 75, 0.2) !important;
        }
        button[kind="primary"]:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 15px rgba(255, 75, 75, 0.35) !important;
        }
        
        /* Checkboxes (Wrapping them like little smart tags) */
        .stCheckbox {
            padding: 4px 8px;
            border-radius: 8px;
            transition: all 0.2s;
        }
        .stCheckbox:hover {
            background-color: #f1f5f9;
        }
        
        /* Beautiful Popups / Dialogs */
        div[data-testid="stDialog"] {
            border-radius: 24px !important;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1) !important;
        }
        
        /* Comment Cards (Glassmorphism inspired) */
        div[data-testid="stVerticalBlock"] > div[style*="border-color:"] {
            border: none !important;
            background: white !important;
            border-radius: 16px !important;
            box-shadow: 0 8px 24px rgba(149, 157, 165, 0.08) !important;
            padding: 1.5rem !important;
            margin-bottom: 1rem !important;
        }
        
        /* Sidebar styling */
        section[data-testid="stSidebar"] {
            background-color: #ffffff;
            border-right: 1px solid #f1f5f9;
        }
        
        /* Hide Default Streamlit Clutter (Headers, Footers) */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        
    </style>
    """, unsafe_allow_html=True)

    main_ws, comment_ws, checklist_ws = get_worksheets()
    if not main_ws or not comment_ws or not checklist_ws:
        st.info("請確認 st.secrets 中的 spreadsheet_id 與 Google Service Account 已正確設定。")
        st.stop()

    # Sidebar: Data Management
    with st.sidebar:
        st.header("⚙️ 系統管理")
        uploaded_file = st.file_uploader("上傳更新 Excel 檔案", type=["xlsx", "xls"])
        if uploaded_file:
            if st.button("確認更新至雲端"):
                # Read everything as string to prevent Excel from dropping leading zeros
                new_df = pd.read_excel(uploaded_file, dtype=str)
                
                # Fix phone numbers if they lost their leading 0 in Excel
                phone_keywords = ["手機", "電話", "phone", "聯絡"]
                for col in new_df.columns:
                    if any(k in str(col).lower() for k in phone_keywords):
                        def fix_phone(val):
                            if pd.isna(val) or val == "": return ""
                            val_str = str(val).split('.')[0]
                            digits = ''.join(filter(str.isdigit, val_str))
                            if len(digits) == 9 and digits.startswith('9'):
                                return '0' + digits
                            return val_str
                        new_df[col] = new_df[col].apply(fix_phone)
                
                # Update MainData
                cols_to_drop_keywords = ['狀況', '案件類型', '續約狀態', '下次續約日期', '派案人員', '接案人員', '營業處', '代管人員', '媒合編號', '備註']
                cols_actual_drop = [c for c in new_df.columns if any(k in str(c) for k in cols_to_drop_keywords)]
                if cols_actual_drop:
                    new_df = new_df.drop(columns=cols_actual_drop)

                # Generate unique identifier for merging
                if '物件縣市' in new_df.columns and '租賃地址' in new_df.columns:
                    new_df['關聯物件地址'] = new_df['物件縣市'].astype(str) + new_df['租賃地址'].astype(str)

                # Merge with existing data
                existing_df = st.session_state.df_main.copy()
                if not existing_df.empty and '關聯物件地址' in existing_df.columns and '關聯物件地址' in new_df.columns:
                    # Concat and drop duplicates, keeping the new upload's data
                    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                    combined_df = combined_df.drop_duplicates(subset=['關聯物件地址'], keep='last')
                else:
                    combined_df = new_df

                save_to_main(main_ws, combined_df)
                st.session_state.df_main = combined_df
                st.success("資料已更新！之前的案件已保留，並且新資料已成功匯入！")
                st.rerun()
        
        st.divider()
        if st.button("🔄 從雲端重新讀取資料", use_container_width=True):
            with st.spinner("下載最新資料中..."):
                st.session_state.df_main = load_main_data(main_ws)
                st.session_state.df_comments = load_comments(comment_ws)
                st.session_state.df_checklist = load_checklist(checklist_ws)
            st.toast("✅ 資料已同步")
            st.rerun()

    # Load Data seamlessly from Session State
    if "df_main" not in st.session_state:
        with st.spinner("初次載入資料中..."):
            st.session_state.df_main = load_main_data(main_ws)
            st.session_state.df_comments = load_comments(comment_ws)
            st.session_state.df_checklist = load_checklist(checklist_ws)

    df_main = st.session_state.df_main.copy()
    df_comments = st.session_state.df_comments.copy()
    df_checklist = st.session_state.df_checklist.copy()

    if df_main.empty:
        st.warning("目前 MainData 中無資料，請先上傳 Excel 檔案。")
        st.stop()

    # Identification mapping for "Expiry Date" and "Address"
    possible_expiry = ["到期日", "租約迄", "結束日期", "Expiry"]
    possible_address = ["地址", "標的名稱", "Address"]
    
    auto_expiry_col = next((c for c in df_main.columns if any(p in c for p in possible_expiry)), df_main.columns[0])
    auto_address_col = next((c for c in df_main.columns if any(p in c for p in possible_address)), df_main.columns[0])

    address_col = auto_address_col # Force use of auto-detected or first column for address

    with st.sidebar:
        st.divider()
        st.subheader("📍 欄位對應設定")
        expiry_col = st.selectbox("請選擇「到期日」欄位", options=df_main.columns, index=df_main.columns.get_loc(auto_expiry_col))

    # Month Filter
    # Excel sometimes stores dates as integers (serial numbers) or strings. 
    # We try to convert safely.
    import re
    def parse_excel_date(val):
        if pd.isna(val) or val == "":
            return pd.NaT
        try:
            # If it's a number (Excel serial date), convert it
            # Excel epoch is 1899-12-30
            val_float = float(val)
            if val_float > 10000: # Plausible Excel date range
                return pd.Timedelta(days=val_float) + pd.Timestamp('1899-12-30')
        except ValueError:
            pass
        
        # Check for ROC year format (e.g., 113/05/20 or 115-06-30)
        if isinstance(val, str):
            # Also handle potentially dots as separators e.g. 115.06.30
            val_str = val.replace('.', '/')
            match = re.match(r'^(\d{2,3})[/-](\d{1,2})[/-](\d{1,2})$', val_str.strip())
            if match:
                y, m, d = match.groups()
                y = int(y)
                if y < 200: # Likely ROC year, add 1911
                    return pd.to_datetime(f"{y+1911}-{m}-{d}", errors='coerce')

        # Otherwise parsing as standard string
        return pd.to_datetime(val, errors='coerce')

    df_main[expiry_col] = df_main[expiry_col].apply(parse_excel_date)
    
    # Filter out rows where date parsing failed for the purpose of the month filter
    valid_dates_df = df_main.dropna(subset=[expiry_col])
    
    if valid_dates_df.empty:
        st.error(f"所選的「{expiry_col}」欄位無法辨識為日期格式。請確認該欄位內容，或在側邊欄選擇正確的日期欄位。")
        st.info("提示：日期格式建議為 2024/01/01 或 2024-01-01。")
        
        # --- DEBUG OUTPUT ---
        st.warning("【系統除錯資訊】以下是系統讀取到的原始資料前 5 筆，請檢查格式是否特殊（例如民國年 113/01/01 或含有特殊符號）：")
        # Load the raw data again just to show the raw strings
        raw_df = load_main_data(main_ws)
        st.write(raw_df[expiry_col].head(5).tolist())
        # --------------------
        
        st.stop()

    df_main['YearMonth'] = df_main[expiry_col].dt.strftime('%Y-%m')
    df_main[expiry_col] = df_main[expiry_col].dt.strftime('%Y-%m-%d')
    
    # Merge df_main with df_checklist for unified filtering
    display_df = df_main.copy()
    if not display_df.empty:
        if '狀態' in display_df.columns:
            display_df = display_df.drop(columns=['狀態'])
        if '已報業績' in display_df.columns:
            display_df = display_df.drop(columns=['已報業績'])
            
        chk_for_merge = df_checklist[['關聯物件地址', '狀態', '已報業績']].copy()
        # Rename the checklist column to match df_main[address_col] so we can merge perfectly on the exact string shown in the UI
        chk_for_merge = chk_for_merge.rename(columns={'關聯物件地址': address_col})
        display_df = display_df.merge(chk_for_merge, on=address_col, how='left')
        display_df['狀態'] = display_df['狀態'].fillna('未送預審')
        display_df['已報業績'] = display_df['已報業績'].fillna('False')
    
    # Search and Filter Dashboard
    st.markdown("### 🔍 搜尋與篩選")
    search_query = st.text_input("輸入物件地址進行搜尋 (不受月份限制)", placeholder="例如：金城路三段")
    
    f_col1, f_col2, f_col3 = st.columns(3)
    available_months = sorted(display_df['YearMonth'].dropna().unique()) if 'YearMonth' in display_df.columns else []
    selected_months = f_col1.multiselect("📅 選擇到期月份 (若搜尋地址則忽略此項)", options=available_months)
    
    status_options = ["未送預審", "預審中", "預審通過", "簽約中", "已完成", "不續約"]
    selected_statuses = f_col2.multiselect("🏷️ 篩選狀態", options=status_options)
    
    perf_option = f_col3.selectbox("💰 篩選已報業績", options=["所有", "已報", "未報"], index=0)

    # Quick Hide features for flexibility
    st.markdown("🎯 **快速隱藏選項**")
    t1, t2, t3 = st.columns(3)
    hide_completed = t1.checkbox("🙈 隱藏「已完成」", value=False)
    hide_no_renew = t2.checkbox("🙈 隱藏「不續約」", value=False)
    hide_perf_done = t3.checkbox("🙈 隱藏「已報業績」", value=False)

    filtered_df = display_df.copy()
    if search_query:
        # Ignore month filter if searching
        filtered_df = filtered_df[filtered_df[address_col].astype(str).str.contains(search_query, na=False, case=False)]
    else:
        # Require month selection if no search query
        if not selected_months:
            st.info("請於上方選取欲查看的月份，或直接搜尋地址。")
            st.stop()
        filtered_df = filtered_df[filtered_df['YearMonth'].isin(selected_months)]
        
    if selected_statuses:
        filtered_df = filtered_df[filtered_df['狀態'].isin(selected_statuses)]
        
    if perf_option == "已報":
        filtered_df = filtered_df[filtered_df['已報業績'] == 'True']
    elif perf_option == "未報":
        filtered_df = filtered_df[filtered_df['已報業績'] != 'True']
        
    # Apply Quick Hide rules
    if hide_completed:
        filtered_df = filtered_df[filtered_df['狀態'] != '已完成']
    if hide_no_renew:
        filtered_df = filtered_df[filtered_df['狀態'] != '不續約']
    if hide_perf_done:
        filtered_df = filtered_df[filtered_df['已報業績'] != 'True']

    with st.sidebar:
        st.divider()
        st.subheader("📋 資料匯出與分群")
        if st.button("📍 顯示地址分群 (台北/新北)", use_container_width=True):
            show_grouped_addresses(filtered_df, address_col)

    # Dynamic Field Selector
    all_fields = [c for c in df_main.columns if c != 'YearMonth']
    # Removed address_col from the default display fields to avoid redundancy
    display_fields = st.multiselect("📊 選擇顯示欄位 (卡片主要內容)", options=all_fields, default=[expiry_col])

    other_fields = [f for f in all_fields if f not in display_fields]

    # --- Table Layout ---
    c_head1, c_head2 = st.columns([8, 2])
    c_head1.markdown("### 📋 案件列表")
    if c_head2.button("➕ 新增物件", type="primary", use_container_width=True):
        show_add_property_dialog(df_main, main_ws, address_col, expiry_col, display_fields, other_fields)
    total_count = len(filtered_df)
    st.caption(f"共找到 {total_count} 筆符合條件的案件")
    
    if total_count == 0:
        st.stop()

    # Table Header
    with st.container():
        h1, h2, h3, h4, h5, h6, h7, h8 = st.columns([2.5, 0.8, 1.2, 0.8, 0.8, 1.5, 0.8, 1.0])
        h1.markdown("**物件狀態及地址 (點擊查看明細)**")
        h2.markdown("**房東**")
        h3.markdown("**房客**")
        h4.markdown("**物件**")
        h5.markdown("**安全**")
        h6.markdown("**狀態**")
        h7.markdown("**業績**")
        h8.markdown("**雲端同步**")
    
    st.divider()

    # Table Rows
    def update_checklist(addr, field, key):
        if checklist_ws is None: return
        new_val = st.session_state[key]
        full_chk = st.session_state.df_checklist.copy()
        idx_list = full_chk.index[full_chk['關聯物件地址'] == addr].tolist()
        if not idx_list:
            new_row = {"關聯物件地址": addr, "房東資料": "", "房客資料": "", "物件資料": "", "安全檢核": "", "證件期限檢核": "", "狀態": "未送預審", "已報業績": "False"}
            new_row[field] = str(new_val)
            full_chk = pd.concat([full_chk, pd.DataFrame([new_row])], ignore_index=True)
        else:
            full_chk.at[idx_list[0], field] = str(new_val)
            
        st.session_state.df_checklist = full_chk

    status_options = ["未送預審", "預審中", "預審通過", "簽約中", "已完成", "不續約"]

    for index, row in filtered_df.iterrows():
        item_address = str(row[address_col])
        
        tenant_val = ""
        for c in row.index:
            if "身分" in c:
                tenant_val = str(row[c])
                break
        
        is_class_1 = "1" in tenant_val or "第一" in tenant_val
        is_class_2 = "2" in tenant_val or "第二" in tenant_val
        needs_expiry_check = is_class_1 or is_class_2
        expiry_label = "更新證件期限" if is_class_2 else "證件期限"
        
        chk_row_idx = df_checklist.index[df_checklist['關聯物件地址'] == item_address].tolist()
        if chk_row_idx:
            chk_data = df_checklist.iloc[chk_row_idx[0]]
            chk1_val = str(chk_data.get("房東資料", "")) == "True"
            chk2_val = str(chk_data.get("房客資料", "")) == "True"
            chk3_val = str(chk_data.get("物件資料", "")) == "True"
            chk4_val = str(chk_data.get("安全檢核", "")) == "True"
            chk5_val = str(chk_data.get("證件期限檢核", "")) == "True"
            chk_status_val = str(chk_data.get("狀態", "未送預審"))
            chk_perf_val = str(chk_data.get("已報業績", "")) == "True"
        else:
            chk1_val, chk2_val, chk3_val, chk4_val, chk5_val = False, False, False, False, False
            chk_status_val = "未送預審"
            chk_perf_val = False

        c1, c2, c3, c4, c5, c6, c7, c8 = st.columns([2.5, 0.8, 1.2, 0.8, 0.8, 1.5, 0.8, 1.0])
        
        # Address Button
        current_status = chk_status_val if chk_status_val in status_options else "未送預審"
        emoji = get_status_emoji(current_status)
        if c1.button(f"{emoji} 🏠 {item_address}", key=f"btn_{item_address}", use_container_width=True):
            show_property_details(row, address_col, display_fields, other_fields, comment_ws, checklist_ws, main_ws)
        
        # Checkboxes
        c2.checkbox("房東", value=chk1_val, key=f"tbl_chk1_{item_address}", on_change=update_checklist, args=(item_address, "房東資料", f"tbl_chk1_{item_address}"))
        
        # Tenant Info + Expiry Check
        c3.checkbox("房客", value=chk2_val, key=f"tbl_chk2_{item_address}", on_change=update_checklist, args=(item_address, "房客資料", f"tbl_chk2_{item_address}"))
        if needs_expiry_check:
            c3.checkbox(f"🚨{expiry_label}", value=chk5_val, key=f"tbl_chk5_{item_address}", on_change=update_checklist, args=(item_address, "證件期限檢核", f"tbl_chk5_{item_address}"))
            
        c4.checkbox("物件", value=chk3_val, key=f"tbl_chk3_{item_address}", on_change=update_checklist, args=(item_address, "物件資料", f"tbl_chk3_{item_address}"))
        c5.checkbox("安全", value=chk4_val, key=f"tbl_chk4_{item_address}", on_change=update_checklist, args=(item_address, "安全檢核", f"tbl_chk4_{item_address}"))
        
        # Status Dropdown
        current_idx = status_options.index(current_status)
        c6.selectbox(" ", options=status_options, index=current_idx, key=f"tbl_status_{item_address}", label_visibility="collapsed", on_change=update_checklist, args=(item_address, "狀態", f"tbl_status_{item_address}"))
        
        # Performance Checkbox
        c7.checkbox("已報", value=chk_perf_val, key=f"tbl_perf_{item_address}", on_change=update_checklist, args=(item_address, "已報業績", f"tbl_perf_{item_address}"))
        
        # Cloud Sync Button
        if c8.button("💾 儲存", key=f"tbl_save_{item_address}", help="將此物件的狀態與留言儲存至雲端", use_container_width=True):
            with st.spinner("同步至雲端中..."):
                f_chk = st.session_state.df_checklist
                if checklist_ws is not None:
                    checklist_ws.clear()
                    checklist_ws.update([f_chk.columns.values.tolist()] + f_chk.values.tolist())
                f_com = st.session_state.df_comments
                if comment_ws is not None:
                    comment_ws.clear()
                    comment_ws.update([f_com.columns.values.tolist()] + f_com.values.tolist())
            st.toast(f"✅ {item_address} 儲存成功！")
        
        # Historical comments thread inline
        with st.container():
            render_inline_comments(item_address, comment_ws, is_dialog=False)
            
        st.divider()


if __name__ == "__main__":
    main()
