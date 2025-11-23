import streamlit as st
from docxtpl import DocxTemplate
import io
from datetime import datetime, date

# --- 1. 頁面設定 (必須在第一行) ---
st.set_page_config(
    page_title="東海大學實習合約系統", 
    page_icon="🎓", 
    layout="centered", # 改回 centered 讓視線更集中，不發散
    initial_sidebar_state="collapsed"
)

# --- 2. CSS 深度魔改 (V4: 極簡商業風格) ---
st.markdown("""
    <style>
    /* 全域字體設定 */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Noto Sans TC', sans-serif;
        color: #2c3e50;
        background-color: #f8f9fa; /* 柔和灰背景 */
    }

    /* 隱藏 Streamlit 原生元素 */
    #MainMenu, footer, header {visibility: hidden;}
    .stDeployButton {display:none;}

    /* === 頂部 Hero 區塊 === */
    .hero-header {
        background: linear-gradient(135deg, #002E5D 0%, #001a35 100%);
        padding: 2rem 1rem;
        margin: -5rem -5rem 2rem -5rem; /* 抵銷 Streamlit 預設 padding */
        text-align: center;
        color: white;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    .hero-title {
        font-size: 1.8rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        letter-spacing: 2px;
    }
    .hero-subtitle {
        color: #C6A87C;
        font-size: 0.9rem;
        font-weight: 500;
        opacity: 0.9;
    }

    /* === 卡片式容器設計 (關鍵) === */
    .stVerticalBlock > div > [data-testid="stVerticalBlock"] {
        background-color: white;
        padding: 2rem;
        border-radius: 16px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.05); /* 浮起效果 */
        border: 1px solid #edf2f7;
        margin-bottom: 1.5rem;
    }

    /* === 輸入框美化 === */
    /* 去除灰底，改用現代化白底+邊框 */
    .stTextInput input, .stNumberInput input, .stSelectbox div[data-baseweb="select"] > div, .stDateInput input, .stTimeInput input {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 10px !important;
        padding: 10px 12px !important;
        font-size: 1rem !important;
        transition: all 0.2s;
        box-shadow: none !important;
    }
    
    /* 輸入框 Focus 效果 */
    .stTextInput input:focus, .stNumberInput input:focus {
        border-color: #002E5D !important;
        box-shadow: 0 0 0 3px rgba(0, 46, 93, 0.1) !important;
    }

    /* 標題優化 */
    h3 {
        color: #002E5D !important;
        font-size: 1.3rem !important;
        font-weight: 700 !important;
        border-bottom: 2px solid #f1f5f9;
        padding-bottom: 10px;
        margin-top: 0 !important;
        margin-bottom: 20px !important;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    /* === 按鈕美化 === */
    .stButton > button {
        width: 100%;
        background: linear-gradient(90deg, #002E5D 0%, #004080 100%);
        color: white;
        border: none;
        padding: 0.8rem 2rem;
        border-radius: 50px; /* 膠囊狀按鈕 */
        font-size: 1.1rem;
        font-weight: 700;
        box-shadow: 0 4px 15px rgba(0, 46, 93, 0.3);
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0, 46, 93, 0.4);
        color: white !important;
    }

    /* 調整 Radio Button 選項間距 */
    div[role="radiogroup"] {
        gap: 1rem;
    }
    </style>
    
    <!-- Hero Header -->
    <div class="hero-header">
        <div class="hero-title">東海大學</div>
        <div class="hero-subtitle">學生校外實習合約產生系統</div>
    </div>
""", unsafe_allow_html=True)

# 準備 Word 變數容器
context = {}

# ==========================================
# 區塊 1：實習機構資料
# ==========================================
st.markdown("<h3>🏢 實習機構資料</h3>", unsafe_allow_html=True)

col1, col2 = st.columns([2, 1])
with col1:
    company_name = st.text_input("機構全銜 (法定名稱)", placeholder="例：國泰世華商業銀行股份有限公司")
with col2:
    company_tax_id = st.text_input("統一編號")

col3, col4 = st.columns(2)
with col3:
    company_rep = st.text_input("代表人姓名")
with col4:
    company_title = st.text_input("代表人職稱", value="負責人")

reg_address = st.text_input("公司登記地址")

# 使用 Expander 收納非必要選項，保持介面整潔
with st.expander("📍 實習地點與登記地址不同？(如派駐分公司)"):
    st.info("💡 系統將自動合併顯示：登記地址 (實習地點：分公司 - 地址)")
    b_col1, b_col2 = st.columns([1, 2])
    with b_col1:
        branch_name = st.text_input("實習單位/分公司名稱", placeholder="例：西屯分公司")
    with b_col2:
        branch_address = st.text_input("實際實習地址")
    
    if branch_name and branch_address:
        final_address = f"{reg_address} (實習地點：{branch_name} - {branch_address})"
    else:
        final_address = reg_address

# ==========================================
# 區塊 2：學生資料
# ==========================================
st.markdown("<h3>🧑‍🎓 實習學生資料</h3>", unsafe_allow_html=True)

student_count = st.radio("本合約學生人數", [1, 2, 3], horizontal=True)

student_list = []
for i in range(student_count):
    st.caption(f"第 {i+1} 位學生")
    s_col1, s_col2 = st.columns([1, 1])
    with s_col1:
        s_name = st.text_input(f"姓名", key=f"s_name_{i}", label_visibility="collapsed", placeholder="學生姓名")
    with s_col2:
        s_id = st.text_input(f"系級 / 學號", key=f"s_id_{i}", label_visibility="collapsed", placeholder="系級 / 學號")
    student_list.append({'name': s_name, 'id': s_id})

# 補足空位
while len(student_list) < 3:
    student_list.append({'name': "", 'id': ""})

# ==========================================
# 區塊 3：實習條件 (核心體驗優化)
# ==========================================
st.markdown("<h3>📝 實習條件與類型</h3>", unsafe_allow_html=True)

# 1. 類型選擇 (使用 Columns 讓選項更明顯)
contract_type = st.radio(
    "請選擇實習類型",
    ("一般型 (學習型)", "工作型 (勞資型)"),
    horizontal=True,
    label_visibility="collapsed"
)

# 2. 日期選擇 (使用日曆選單，後台轉民國年)
st.write("實習期間")
d_col1, d_col2 = st.columns(2)
with d_col1:
    s_date = st.date_input("開始日期", value=date(2024, 7, 1))
with d_col2:
    e_date = st.date_input("結束日期", value=date(2025, 6, 30))

# 自動計算民國年
s_y, s_m, s_d = s_date.year - 1911, s_date.month, s_date.day
e_y, e_m, e_d = e_date.year - 1911, e_date.month, e_date.day

# 3. 時間選擇
st.write("每日實習時間")
t_col1, t_col2, t_col3 = st.columns([1, 1, 1])
with t_col1:
    daily_start = st.time_input("開始", value=datetime.strptime("09:00", "%H:%M"))
with t_col2:
    daily_end = st.time_input("結束", value=datetime.strptime("18:00", "%H:%M"))
with t_col3:
    daily_hours = st.number_input("每日時數", value=8.0, step=0.5)

# ==========================================
# 區塊 4：待遇與福利 (動態色塊)
# ==========================================
st.markdown("<h3>💰 待遇與福利</h3>", unsafe_allow_html=True)

# 初始化變數
context.update({
    'type_learn_check': '□', 'type_work_check': '□',
    'chk_pay_none': '□', 'chk_pay_scholar': '□', 'chk_pay_allowance': '□',
    'pay_learn_amount': "", 'pay_work_amount': ""
})

if contract_type == "一般型 (學習型)":
    st.success("✅ **學習型**：單純學習訓練，無僱傭關係。")
    context['type_learn_check'] = '☑'
    
    pay_opt = st.radio("給付類型", ["無", "獎學金", "實習津貼"], horizontal=True)
    
    if pay_opt == "無":
        context['chk_pay_none'] = '☑'
        context['pay_learn_amount'] = "0"
    elif pay_opt == "獎學金":
        context['chk_pay_scholar'] = '☑'
        amt = st.number_input("獎學金金額 (元)", min_value=0, step=1000)
        context['pay_learn_amount'] = f"{amt:,}"
    else:
        context['chk_pay_allowance'] = '☑'
        amt = st.number_input("津貼金額 (元)", min_value=0, step=1000)
        context['pay_learn_amount'] = f"{amt:,}"
        
else: 
    st.warning("⚠️ **勞資型**：具僱傭關係，需投保勞健保。")
    context['type_work_check'] = '☑'
    
    pay_work_amt = st.number_input("每月薪資 (元)", min_value=27470, step=100, help="請確認符合當年度基本工資")
    context['pay_work_amount'] = f"{pay_work_amt:,}"

st.markdown("---")

# 福利 Helper
def welfare_ui(title, key_prefix, unit):
    c1, c2 = st.columns([1, 2])
    with c1:
        st.write(f"**{title}**")
    with c2:
        opt = st.selectbox(f"{title}選項", ["無", "免費提供", "付費提供"], key=key_prefix, label_visibility="collapsed")
        cost_txt = ""
        checks = {f'chk_{key_prefix}_none': '□', f'chk_{key_prefix}_free': '□', f'chk_{key_prefix}_paid': '□'}
        
        if opt == "無": checks[f'chk_{key_prefix}_none'] = '☑'
        elif opt == "免費提供": checks[f'chk_{key_prefix}_free'] = '☑'
        else:
            checks[f'chk_{key_prefix}_paid'] = '☑'
            val = st.number_input(f"費用 ({unit})", min_value=0, step=100, key=f"{key_prefix}_cost")
            cost_txt = f"{val:,}"
    return checks, cost_txt

d_checks, d_cost = welfare_ui("住宿", "dorm", "元/月")
f_checks, f_cost = welfare_ui("膳食", "food", "元/餐")

# 交通特別處理
c1, c2 = st.columns([1, 2])
with c1: st.write("**交通**")
with c2:
    t_opt = st.selectbox("交通選項", ["無", "免費提供", "付費提供"], key="trans", label_visibility="collapsed")
    t_checks = {'chk_trans_none': '□', 'chk_trans_free': '□', 'chk_trans_paid': '□'}
    t_cost = ""
    if t_opt == "無": t_checks['chk_trans_none'] = '☑'
    elif t_opt == "免費提供": t_checks['chk_trans_free'] = '☑'
    else: 
        t_checks['chk_trans_paid'] = '☑'
        val = st.number_input("交通費用/津貼 (元/月)", min_value=0, step=100)
        t_cost = f"{val:,}"

context.update(d_checks); context.update({'dorm_cost': d_cost})
context.update(f_checks); context.update({'food_cost': f_cost})
context.update(t_checks); context.update({'trans_cost': t_cost})

# ==========================================
# 底部按鈕區
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
generate_btn = st.button("🚀 產生合約 (Word)", type="primary")

if generate_btn:
    if not company_name or not student_list[0]['name']:
        st.error("❌ 請填寫「機構名稱」與「第一位學生姓名」")
    else:
        # 填入變數
        context.update({
            'company_name': company_name,
            'company_tax_id': company_tax_id,
            'company_rep': company_rep,
            'company_title': company_title,
            'company_address': final_address if 'final_address' in locals() else reg_address,
            's1_name': student_list[0]['name'], 's1_id': student_list[0]['id'],
            's2_name': student_list[1]['name'], 's2_id': student_list[1]['id'],
            's3_name': student_list[2]['name'], 's3_id': student_list[2]['id'],
            'student_name': student_list[0]['name'] + (" 等" if student_count > 1 else ""),
            's_y': s_y, 's_m': s_m, 's_d': s_d,
            'e_y': e_y, 'e_m': e_m, 'e_d': e_d,
            'daily_start': daily_start.strftime("%H:%M"),
            'daily_end': daily_end.strftime("%H:%M"),
            'daily_hours': daily_hours,
        })
        
        try:
            doc = DocxTemplate("template.docx")
            doc.render(context)
            bio = io.BytesIO()
            doc.save(bio)
            
            st.balloons()
            st.success("✅ 合約已建立！")
            st.download_button(
                label="📥 下載 Word 檔",
                data=bio.getvalue(),
                file_name=f"東海大學實習合約_{student_list[0]['name']}_{company_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"❌ 錯誤：{e}")
