import streamlit as st
from docxtpl import DocxTemplate
import io
from datetime import datetime, date

# --- 1. 頁面設定 (必須在第一行) ---
st.set_page_config(
    page_title="東海大學實習合約產生系統", 
    page_icon="🎓", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

# --- 2. CSS 深度美化 (打造像原生網頁的質感) ---
st.markdown("""
    <style>
    /* 引入字體 */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Noto Sans TC', sans-serif;
        color: #333333;
    }

    /* === 頂部導覽列樣式 === */
    .thu-header {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        background-color: #002E5D; /* 東海深藍 */
        color: white;
        padding: 0.8rem 2rem;
        z-index: 999999; /* 確保在最上層 */
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    .thu-header h1 {
        margin: 0;
        font-size: 1.4rem;
        color: white;
        font-weight: 700;
        letter-spacing: 1px;
    }
    
    .thu-header span {
        color: #C6A87C; /* 東海金 */
        font-size: 0.9rem;
        margin-left: 12px;
        font-weight: 500;
    }

    /* 隱藏 Streamlit 預設的漢堡選單與 Footer，讓介面更乾淨 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;} /* 隱藏原本的頂部白條 */

    /* 調整主內容往下，避免被導覽列遮住 */
    .block-container {
        padding-top: 6rem; 
        padding-bottom: 5rem;
        max-width: 960px; /* 限制寬度，閱讀體驗更好 */
    }

    /* === 輸入框美化 (關鍵：去灰底，改白底邊框) === */
    .stTextInput input, .stSelectbox div[data-baseweb="select"] > div, .stNumberInput input, .stDateInput input, .stTimeInput input {
        background-color: #ffffff !important; /* 強制白底 */
        border: 1px solid #ced4da !important; /* 灰色細邊框 */
        border-radius: 6px !important;
        color: #495057 !important;
        padding: 0.5rem !important;
    }
    
    /* 輸入框 Focus 狀態 (點擊時變東海藍) */
    .stTextInput input:focus, .stNumberInput input:focus, .stDateInput input:focus {
        border-color: #002E5D !important;
        box-shadow: 0 0 0 3px rgba(0, 46, 93, 0.15) !important;
    }

    /* 步驟條樣式優化 */
    .step-container {
        display: flex;
        justify-content: space-between;
        margin-bottom: 2.5rem;
        padding: 0 3rem;
        position: relative;
    }
    .step-item {
        display: flex;
        flex-direction: column;
        align-items: center;
        position: relative;
        flex: 1;
        z-index: 2;
    }
    .step-circle {
        width: 32px;
        height: 32px;
        background-color: #002E5D;
        color: white;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: bold;
        margin-bottom: 0.5rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    .step-label {
        font-size: 0.85rem;
        font-weight: 700;
        color: #555;
    }
    /* 連接線 */
    .step-line-bg {
        position: absolute;
        top: 16px;
        left: 15%;
        width: 70%;
        height: 2px;
        background-color: #e0e0e0;
        z-index: 1;
    }

    /* === 按鈕美化 === */
    .stButton>button {
        background-color: #002E5D;
        color: white;
        border-radius: 8px;
        font-weight: bold;
        padding: 0.6rem 2rem;
        border: none;
        box-shadow: 0 4px 6px rgba(0, 46, 93, 0.3);
        transition: all 0.3s;
        width: 100%;
        font-size: 1.1rem;
    }
    .stButton>button:hover {
        background-color: #001a35;
        box-shadow: 0 6px 8px rgba(0, 46, 93, 0.4);
        transform: translateY(-2px);
    }
    
    /* 卡片標題裝飾 */
    .section-title {
        color: #002E5D;
        font-size: 1.25rem;
        font-weight: 700;
        margin-bottom: 1.2rem;
        border-left: 6px solid #C6A87C;
        padding-left: 12px;
        line-height: 1.2;
    }
    
    /* 調整 Container 邊框顏色 */
    [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {
        background-color: white;
    }
    </style>
    
    <!-- 頂部導覽列 HTML -->
    <div class="thu-header">
        <div style="display:flex; align-items:center;">
            <h1>東海大學</h1>
            <span>學生校外實習合約系統</span>
        </div>
        <div style="background: rgba(255,255,255,0.15); padding: 4px 12px; border-radius: 20px; font-size: 0.8rem; border: 1px solid rgba(255,255,255,0.2);">
            <span style="color:white; margin:0;">👤 承辦人員模式</span>
        </div>
    </div>
    
    <!-- 步驟條 HTML -->
    <div class="step-container">
        <div class="step-line-bg"></div>
        <div class="step-item">
            <div class="step-circle">1</div>
            <div class="step-label">機構資料</div>
        </div>
        <div class="step-item">
            <div class="step-circle">2</div>
            <div class="step-label">學生資料</div>
        </div>
        <div class="step-item">
            <div class="step-circle">3</div>
            <div class="step-label">預覽與下載</div>
        </div>
    </div>
""", unsafe_allow_html=True)

# 準備 Word 變數容器
context = {}

# ==========================================
# 區塊 1：實習機構資料
# ==========================================
with st.container(border=True):
    st.markdown('<div class="section-title">🏢 乙方：實習機構資料</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    with col1:
        company_name = st.text_input("機構全銜 (法定名稱) *", placeholder="請輸入完整名稱，例：國泰世華商業銀行股份有限公司")
    with col2:
        company_tax_id = st.text_input("統一編號")

    col3, col4 = st.columns(2)
    with col3:
        company_rep = st.text_input("代表人姓名")
    with col4:
        company_title = st.text_input("代表人職稱", value="負責人")
        
    st.markdown("---")
    
    reg_address = st.text_input("公司登記地址")
    
    # 分公司邏輯
    is_branch = st.checkbox("📍 實習地點與登記地址不同 (如派駐分公司)")
    if is_branch:
        st.info("💡 系統將自動合併顯示：登記地址 (實習地點：分公司 - 地址)")
        b_col1, b_col2 = st.columns([1, 2])
        with b_col1:
            branch_name = st.text_input("實習單位/分公司名稱", placeholder="例：西屯分公司")
        with b_col2:
            branch_address = st.text_input("實際實習地址")
        final_address = f"{reg_address} (實習地點：{branch_name} - {branch_address})"
    else:
        final_address = reg_address

# ==========================================
# 區塊 2：學生資料
# ==========================================
with st.container(border=True):
    st.markdown('<div class="section-title">🧑‍🎓 甲方：實習學生資料</div>', unsafe_allow_html=True)
    
    sc_col1, sc_col2 = st.columns([1, 3])
    with sc_col1:
        student_count = st.number_input("本合約學生人數", min_value=1, max_value=3, value=1)
    
    student_list = []
    
    for i in range(student_count):
        st.markdown(f"**第 {i+1} 位學生**")
        s_col1, s_col2 = st.columns(2)
        with s_col1:
            s_name = st.text_input(f"姓名", key=f"s_name_{i}")
        with s_col2:
            s_id = st.text_input(f"系級 / 學號", key=f"s_id_{i}", placeholder="例：國貿四A / s109...")
        student_list.append({'name': s_name, 'id': s_id})
    
    # 補足空位
    while len(student_list) < 3:
        student_list.append({'name': "", 'id': ""})

# ==========================================
# 區塊 3：實習條件
# ==========================================
with st.container(border=True):
    st.markdown('<div class="section-title">📝 實習條件與類型</div>', unsafe_allow_html=True)
    
    type_col1, type_col2 = st.columns(2)
    with type_col1:
        contract_type = st.radio(
            "1. 請選擇實習類型 (將連動條款)",
            ("一般型 (學習型)", "工作型 (勞資型)"),
            horizontal=True
        )

    st.markdown("**2. 實習期間 (民國年)**")
    # 調整欄位比例讓顯示更緊湊
    d_col1, d_col2, d_col3, d_col4 = st.columns([0.1, 1.2, 0.1, 1.2])
    with d_col1:
        st.write("自")
    with d_col2:
        curr_year = datetime.now().year - 1911
        c1, c2, c3 = st.columns(3)
        s_y = c1.number_input("年", 113, 120, curr_year, key="sy")
        s_m = c2.number_input("月", 1, 12, 7, key="sm")
        s_d = c3.number_input("日", 1, 31, 1, key="sd")
    with d_col3:
        st.write("至")
    with d_col4:
        c4, c5, c6 = st.columns(3)
        e_y = c4.number_input("年", 113, 120, curr_year+1, key="ey")
        e_m = c5.number_input("月", 1, 12, 6, key="em")
        e_d = c6.number_input("日", 1, 31, 30, key="ed")

    st.markdown("**3. 每日實習時間**")
    t_col1, t_col2, t_col3 = st.columns(3)
    with t_col1:
        daily_start = st.time_input("開始時間", value=datetime.strptime("09:00", "%H:%M"))
    with t_col2:
        daily_end = st.time_input("結束時間", value=datetime.strptime("18:00", "%H:%M"))
    with t_col3:
        daily_hours = st.number_input("每日共計 (小時)", value=8.0, step=0.5)

# ==========================================
# 區塊 4：待遇與福利
# ==========================================
with st.container(border=True):
    st.markdown('<div class="section-title">💰 待遇與福利</div>', unsafe_allow_html=True)

    context.update({
        'type_learn_check': '□', 'type_work_check': '□',
        'chk_pay_none': '□', 'chk_pay_scholar': '□', 'chk_pay_allowance': '□',
        'pay_learn_amount': "", 'pay_work_amount': ""
    })

    if contract_type == "一般型 (學習型)":
        st.success("✅ **學習型適用**：單純學習訓練，無僱傭關係。每日不得超過 8 小時。")
        context['type_learn_check'] = '☑'
        
        st.markdown("**給付項目 (每月給付總額)**")
        pay_opt = st.radio("給付類型", ["無", "獎學金", "實習津貼"], horizontal=True, label_visibility="collapsed")
        
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
        st.warning("⚠️ **勞資型適用**：具僱傭關係，需投保勞健保。薪資不得低於基本工資。")
        context['type_work_check'] = '☑'
        
        st.markdown("**薪資待遇**")
        pay_work_amt = st.number_input("每月薪資 (元)", min_value=27470, step=100, help="請確認符合當年度基本工資")
        context['pay_work_amount'] = f"{pay_work_amt:,}"

    st.markdown("---")
    
    # 福利 Helper
    def welfare_ui(title, key_prefix, unit):
        st.markdown(f"**{title}**")
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

    w_col1, w_col2, w_col3 = st.columns(3)
    
    with w_col1: d_checks, d_cost = welfare_ui("住宿", "dorm", "元/月")
    with w_col2: f_checks, f_cost = welfare_ui("膳食", "food", "元/餐")
    with w_col3:
        st.markdown("**交通**")
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
generate_btn = st.button("🚀 產生並下載合約文件 (Word)", type="primary")

if generate_btn:
    if not company_name or not student_list[0]['name']:
        st.error("❌ 請檢查「機構名稱」與「第一位學生姓名」是否已填寫。")
    else:
        # 填入變數
        context.update({
            'company_name': company_name,
            'company_tax_id': company_tax_id,
            'company_rep': company_rep,
            'company_title': company_title,
            'company_address': final_address,
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
            st.success("✅ 合約產生成功！請點擊下方按鈕下載。")
            st.download_button(
                label="📥 點此下載 Word 檔",
                data=bio.getvalue(),
                file_name=f"東海大學實習合約_{student_list[0]['name']}_{company_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"❌ 發生錯誤：{e}")
            st.info("請確認 template.docx 是否與程式在同一目錄下。")
