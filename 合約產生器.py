import streamlit as st
from docxtpl import DocxTemplate
import io
from datetime import datetime, date

# --- 1. 頁面設定 ---
st.set_page_config(
    page_title="東海大學實習合約產生系統", 
    page_icon="🎓", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 2. 極致 CSS 注入 (還原 HTML 模板) ---
st.markdown("""
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
    
    /* 全域設定：模仿 Tailwind 的 bg-gray-100 */
    .stApp {
        background-color: #f3f4f6;
        font-family: 'Noto Sans TC', sans-serif;
    }

    /* 隱藏 Streamlit 原生 Header/Footer/Menu */
    header[data-testid="stHeader"] {display: none;}
    footer {display: none;}
    #MainMenu {display: none;}
    
    /* 調整內容區塊邊距，避開我們自製的 Header */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 5rem !important;
        max-width: 1024px !important;
    }

    /* === 自定義 Header === */
    .thu-header {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 70px;
        background-color: #002E5D; /* 東海藍 */
        color: white;
        z-index: 999999;
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 2rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }

    /* === 步驟條樣式 === */
    .step-wrapper {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin: 6rem 0 2rem 0; /* Top margin to clear header */
        padding: 0 2rem;
        position: relative;
    }
    .step-item {
        display: flex;
        flex-direction: column;
        align-items: center;
        z-index: 2;
    }
    .step-circle {
        width: 2rem;
        height: 2rem;
        background-color: #002E5D;
        color: white;
        border-radius: 9999px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    .step-line {
        flex: 1;
        height: 4px;
        background-color: #d1d5db;
        margin: 0 1rem;
        margin-bottom: 1.5rem;
    }

    /* === 卡片容器魔改 (Target Streamlit Containers) === */
    /* 這是最關鍵的部分：把 st.container(border=True) 變成我們設計的卡片 */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: white;
        border-radius: 0.5rem; /* rounded-lg */
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); /* shadow-md */
        border: none !important;
        padding: 0 !important;
        overflow: hidden;
        margin-bottom: 1.5rem;
    }
    
    /* 卡片內容 Padding */
    div[data-testid="stVerticalBlockBorderWrapper"] > div:nth-child(1) > div {
        padding: 1.5rem; 
    }

    /* 卡片標題區塊樣式 (我們用 HTML 寫在 container 第一行) */
    .card-header {
        background-color: #f9fafb; /* bg-gray-50 */
        padding: 1rem 1.5rem;
        border-bottom: 1px solid #e5e7eb;
        color: #002E5D;
        font-weight: 700;
        font-size: 1.125rem;
        display: flex;
        align-items: center;
        margin: -1.5rem -1.5rem 1.5rem -1.5rem; /* 抵銷 container 的 padding */
    }
    .card-header i {
        margin-right: 0.5rem;
    }
    
    /* 頂部邊框顏色 (透過 inline style 控制) */
    .border-top-blue { border-top: 4px solid #002E5D; }
    .border-top-gold { border-top: 4px solid #C6A87C; }

    /* === 輸入框美化 (去灰底，改白底) === */
    .stTextInput input, .stNumberInput input, .stSelectbox div[data-baseweb="select"] > div, .stDateInput input, .stTimeInput input {
        background-color: white !important;
        border: 1px solid #d1d5db !important;
        border-radius: 0.375rem !important;
        color: #374151 !important;
    }
    .stTextInput input:focus, .stNumberInput input:focus {
        border-color: #002E5D !important;
        box-shadow: 0 0 0 2px rgba(0, 46, 93, 0.2) !important;
    }

    /* === 按鈕美化 === */
    .stButton > button {
        background-color: #002E5D;
        color: white;
        border-radius: 0.375rem;
        padding: 0.5rem 2rem;
        border: none;
        font-weight: bold;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);
        width: 100%;
    }
    .stButton > button:hover {
        background-color: #1e3a8a;
        color: white;
    }
    
    /* Footer */
    .thu-footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: #1f2937;
        color: #9ca3af;
        text-align: center;
        padding: 1rem;
        font-size: 0.875rem;
        z-index: 9999;
    }
    </style>

    <!-- 1. 注入 Header -->
    <div class="thu-header">
        <div style="display:flex; align-items:center; gap:0.75rem;">
            <div style="width:2.5rem; height:2.5rem; background:white; border-radius:9999px; display:flex; align-items:center; justify-content:center; color:#002E5D; font-weight:bold; font-size:1.25rem;">
                <i class="fa-solid fa-university"></i>
            </div>
            <div>
                <div style="font-size:1.25rem; font-weight:bold; letter-spacing:0.05em;">東海大學</div>
                <div style="font-size:0.75rem; color:#C6A87C; opacity:0.9;">學生校外實習合約產生系統</div>
            </div>
        </div>
        <div style="display:none; @media (min-width: 768px) {display:block;}">
            <span style="font-size:0.875rem; background:#002E5D; border:1px solid #C6A87C; padding:0.25rem 0.75rem; border-radius:9999px; color:#C6A87C;">
                <i class="fa-solid fa-user-tie"></i> 承辦人員模式
            </span>
        </div>
    </div>

    <!-- 2. 注入步驟條 -->
    <div class="step-wrapper">
        <div class="step-item">
            <div class="step-circle">1</div>
            <div style="font-size:0.875rem; font-weight:500;">機構資料</div>
        </div>
        <div class="step-line"></div>
        <div class="step-item">
            <div class="step-circle">2</div>
            <div style="font-size:0.875rem; font-weight:500;">學生資料</div>
        </div>
        <div class="step-line"></div>
        <div class="step-item">
            <div style="width:2rem; height:2rem; background:#d1d5db; color:#4b5563; border-radius:9999px; display:flex; align-items:center; justify-content:center; font-weight:bold; margin-bottom:0.5rem;">3</div>
            <div style="font-size:0.875rem; font-weight:500;">實習條件</div>
        </div>
    </div>
""", unsafe_allow_html=True)

# 準備 Word 變數容器
context = {}

# ==========================================
# 卡片 1：實習機構資料 (藍色頂邊框)
# ==========================================
with st.container(border=True): # 這裡的 border=True 會被 CSS 攔截並改造成卡片樣式
    # 注入卡片標題 HTML (包含藍色頂邊框 class)
    st.markdown("""
        <div class="card-header border-top-blue">
            <i class="fa-regular fa-building"></i> 乙方：實習機構資料
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.caption("機構全銜 (法定名稱) *")
        company_name = st.text_input("company_name", placeholder="例：國泰世華商業銀行股份有限公司", label_visibility="collapsed")
    with col2:
        st.caption("統一編號")
        company_tax_id = st.text_input("tax_id", label_visibility="collapsed")

    col3, col4 = st.columns(2)
    with col3:
        st.caption("代表人姓名")
        company_rep = st.text_input("rep_name", label_visibility="collapsed")
    with col4:
        st.caption("代表人職稱")
        company_title = st.text_input("rep_title", value="負責人", label_visibility="collapsed")
    
    st.caption("公司登記地址")
    reg_address = st.text_input("reg_addr", label_visibility="collapsed")
    
    # 分公司區塊
    st.markdown("<div style='border-top:1px solid #e5e7eb; margin:1rem 0;'></div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:0.875rem; font-weight:700; color:#6b7280; margin-bottom:0.75rem;'><i class='fa-solid fa-location-dot'></i> 實際實習地點 (若與登記地址不同請填寫)</div>", unsafe_allow_html=True)
    
    b_col1, b_col2 = st.columns([1, 2])
    with b_col1:
        st.caption("實習單位/分公司名稱")
        branch_name = st.text_input("branch_name", placeholder="例：西屯分公司", label_visibility="collapsed")
    with b_col2:
        st.caption("實習地址")
        branch_address = st.text_input("branch_addr", placeholder="例：臺中市西屯區朝富路217號", label_visibility="collapsed")
    
    if branch_name and branch_address:
        final_address = f"{reg_address} (實習地點：{branch_name} - {branch_address})"
    else:
        final_address = reg_address

# ==========================================
# 卡片 2：學生資料 (藍色頂邊框)
# ==========================================
with st.container(border=True):
    st.markdown("""
        <div class="card-header border-top-blue" style="justify-content:space-between;">
            <div><i class="fa-solid fa-user-graduate"></i> 甲方：實習學生資料</div>
        </div>
    """, unsafe_allow_html=True)

    # 學生人數選擇
    st.caption("本合約學生人數")
    student_count = st.radio("count", [1, 2, 3], horizontal=True, label_visibility="collapsed")
    
    student_list = []
    
    # 學生輸入框
    st.markdown("<div style='background:#eff6ff; border:1px solid #dbeafe; padding:1rem; border-radius:0.5rem;'>", unsafe_allow_html=True)
    for i in range(student_count):
        st.markdown(f"<div style='font-size:0.875rem; font-weight:bold; color:#1e40af; margin-bottom:0.5rem;'>學生 {i+1}</div>", unsafe_allow_html=True)
        s_col1, s_col2 = st.columns(2)
        with s_col1:
            s_name = st.text_input(f"姓名", key=f"s_name_{i}", placeholder="姓名", label_visibility="collapsed")
        with s_col2:
            s_id = st.text_input(f"系級/學號", key=f"s_id_{i}", placeholder="系級/學號", label_visibility="collapsed")
        if i < student_count - 1:
            st.markdown("<hr style='margin:0.5rem 0; border-color:#dbeafe;'>", unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # 補足空位
    for i in range(student_count):
        student_list.append({'name': st.session_state[f"s_name_{i}"], 'id': st.session_state[f"s_id_{i}"]})
    while len(student_list) < 3:
        student_list.append({'name': "", 'id': ""})

# ==========================================
# 卡片 3：實習條件 (金色頂邊框)
# ==========================================
with st.container(border=True):
    st.markdown("""
        <div class="card-header border-top-gold">
            <i class="fa-solid fa-briefcase"></i> 實習條件設定
        </div>
    """, unsafe_allow_html=True)

    # 1. 類型
    st.markdown("<div style='font-size:1.125rem; font-weight:bold; margin-bottom:0.75rem;'>1. 請選擇實習類型</div>", unsafe_allow_html=True)
    
    # 模擬 HTML 的大按鈕樣式 (使用 columns + metric 或是單純 radio)
    # 為了功能性，我們還是用 radio，但加上說明
    contract_type = st.radio("type", ("一般型 (學習型)", "工作型 (勞資型)"), horizontal=True, label_visibility="collapsed")
    
    if contract_type == "一般型 (學習型)":
        st.info("💡 單純學習訓練，無僱傭關係。每日不得超過 8 小時。")
    else:
        st.warning("⚠️ 具僱傭關係，適用勞基法。需投保勞健保。")

    # 2. 期間
    st.markdown("<div style='font-size:1.125rem; font-weight:bold; margin:1.5rem 0 0.75rem 0;'>2. 實習期間 (民國年)</div>", unsafe_allow_html=True)
    d_col1, d_col2 = st.columns(2)
    with d_col1:
        st.caption("開始日期")
        s_date = st.date_input("start_date", value=date(2024, 7, 1), label_visibility="collapsed")
    with d_col2:
        st.caption("結束日期")
        e_date = st.date_input("end_date", value=date(2025, 6, 30), label_visibility="collapsed")
    
    s_y, s_m, s_d = s_date.year - 1911, s_date.month, s_date.day
    e_y, e_m, e_d = e_date.year - 1911, e_date.month, e_date.day

    # 3. 時間
    st.markdown("<div style='font-size:1.125rem; font-weight:bold; margin:1.5rem 0 0.75rem 0;'>3. 每日實習時間</div>", unsafe_allow_html=True)
    t_container = st.container()
    with t_container:
        st.markdown("<div style='background:#f9fafb; padding:1rem; border:1px solid #e5e7eb; border-radius:0.375rem;'>", unsafe_allow_html=True)
        tc1, tc2, tc3 = st.columns([1,1,1])
        with tc1:
            st.caption("開始")
            daily_start = st.time_input("t_start", value=datetime.strptime("09:00", "%H:%M"), label_visibility="collapsed")
        with tc2:
            st.caption("結束")
            daily_end = st.time_input("t_end", value=datetime.strptime("18:00", "%H:%M"), label_visibility="collapsed")
        with tc3:
            st.caption("共計 (小時)")
            daily_hours = st.number_input("hours", value=8.0, step=0.5, label_visibility="collapsed")
        st.markdown("</div>", unsafe_allow_html=True)

    # 4. 待遇
    st.markdown("<div style='border-top:1px solid #e5e7eb; margin:1.5rem 0;'></div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:1.125rem; font-weight:bold; margin-bottom:0.75rem;'>4. 實習待遇與給付</div>", unsafe_allow_html=True)

    # 變數初始化
    context.update({
        'type_learn_check': '□', 'type_work_check': '□',
        'chk_pay_none': '□', 'chk_pay_scholar': '□', 'chk_pay_allowance': '□',
        'pay_learn_amount': "", 'pay_work_amount': ""
    })

    if contract_type == "一般型 (學習型)":
        context['type_learn_check'] = '☑'
        st.markdown("<div style='background:#f0fdf4; border:1px solid #bbf7d0; padding:1rem; border-radius:0.5rem; margin-bottom:1rem;'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#166534; font-weight:bold; font-size:0.875rem; margin-bottom:0.5rem;'>給付項目 (每月給付總額)</div>", unsafe_allow_html=True)
        
        pay_opt = st.radio("pay_opt_learn", ["無", "獎學金", "實習津貼"], horizontal=True, label_visibility="collapsed")
        
        if pay_opt != "無":
            st.caption("金額 (新台幣)")
            amt = st.number_input("amount_learn", min_value=0, step=1000, label_visibility="collapsed")
            context['pay_learn_amount'] = f"{amt:,}"
            if pay_opt == "獎學金": context['chk_pay_scholar'] = '☑'
            else: context['chk_pay_allowance'] = '☑'
        else:
            context['chk_pay_none'] = '☑'
            context['pay_learn_amount'] = "0"
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        context['type_work_check'] = '☑'
        st.markdown("<div style='background:#fff7ed; border:1px solid #fed7aa; padding:1rem; border-radius:0.5rem; margin-bottom:1rem;'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#9a3412; font-weight:bold; font-size:0.875rem; margin-bottom:0.5rem;'>薪資金額 (不得低於基本工資)</div>", unsafe_allow_html=True)
        
        pay_work_amt = st.number_input("amount_work", value=27470, step=100, label_visibility="collapsed")
        context['pay_work_amount'] = f"{pay_work_amt:,}"
        st.markdown("</div>", unsafe_allow_html=True)

    # 5. 福利
    st.markdown("<div style='border-top:1px solid #e5e7eb; margin:1.5rem 0;'></div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:1.125rem; font-weight:bold; margin-bottom:0.75rem;'>5. 福利項目</div>", unsafe_allow_html=True)
    
    w1, w2, w3 = st.columns(3)
    
    def welfare_widget(col, title, key_prefix, unit):
        with col:
            st.caption(title)
            opt = st.selectbox(f"{title}選項", ["無", "免費提供", "付費提供", "交通津貼"] if key_prefix=='trans' else ["無", "免費提供", "付費提供"], key=key_prefix, label_visibility="collapsed")
            
            checks = {f'chk_{key_prefix}_none': '□', f'chk_{key_prefix}_free': '□', f'chk_{key_prefix}_paid': '□'}
            cost_txt = ""
            
            if opt == "無": checks[f'chk_{key_prefix}_none'] = '☑'
            elif opt == "免費提供": checks[f'chk_{key_prefix}_free'] = '☑'
            else:
                checks[f'chk_{key_prefix}_paid'] = '☑'
                # 動態顯示輸入框
                val = st.number_input(f"費用", placeholder=f"{unit}", min_value=0, step=100, key=f"{key_prefix}_cost", label_visibility="collapsed")
                cost_txt = f"{val:,}"
            return checks, cost_txt

    d_checks, d_cost = welfare_widget(w1, "住宿", "dorm", "元/月")
    f_checks, f_cost = welfare_widget(w2, "膳食", "food", "元/餐")
    
    # 交通自己寫因為多一個選項
    with w3:
        st.caption("交通")
        t_opt = st.selectbox("trans_opt", ["無", "免費提供", "付費提供", "交通津貼"], label_visibility="collapsed")
        t_checks = {'chk_trans_none': '□', 'chk_trans_free': '□', 'chk_trans_paid': '□'}
        t_cost = ""
        if t_opt == "無": t_checks['chk_trans_none'] = '☑'
        elif t_opt == "免費提供": t_checks['chk_trans_free'] = '☑'
        else:
            t_checks['chk_trans_paid'] = '☑'
            val = st.number_input("trans_val", min_value=0, step=100, label_visibility="collapsed")
            t_cost = f"{val:,}"

    context.update(d_checks); context.update({'dorm_cost': d_cost})
    context.update(f_checks); context.update({'food_cost': f_cost})
    context.update(t_checks); context.update({'trans_cost': t_cost})

# ==========================================
# 底部按鈕
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
col_btn_1, col_btn_2 = st.columns([1, 2])
with col_btn_1:
    if st.button("清除重填"):
        st.rerun()
with col_btn_2:
    generate_btn = st.button("🚀 產生合約文件 (Word)", type="primary")

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
            st.success("✅ 合約產生成功！")
            st.download_button(
                label="📥 下載 Word 合約檔",
                data=bio.getvalue(),
                file_name=f"東海大學實習合約_{student_list[0]['name']}_{company_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"❌ 錯誤：{e}")

# 注入 Footer
st.markdown("""
    <div class="thu-footer">
        &copy; 2024 東海大學 Tunghai University. All Rights Reserved. | 系統版本 v5.0
    </div>
""", unsafe_allow_html=True)
