import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
from io import BytesIO

st.set_page_config(page_title="Учет закупок кафе", page_icon="📦", layout="wide")

# === ПОДКЛЮЧЕНИЕ К GOOGLE SHEETS ===
@st.cache_resource
def get_sheet():
    try:
        creds_dict = st.secrets["google"]
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet = client.open("Кафе_Закупки").sheet1
        return sheet
    except Exception as e:
        st.error(f"Ошибка подключения: {e}")
        return None

def get_report_sheet():
    try:
        creds_dict = st.secrets["google"]
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        workbook = client.open("Кафе_Закупки")
        try:
            report_sheet = workbook.worksheet("ОТЧЕТЫ")
        except:
            report_sheet = workbook.add_worksheet(title="ОТЧЕТЫ", rows="1000", cols="20")
        return report_sheet
    except Exception as e:
        st.error(f"Ошибка: {e}")
        return None

def load_data():
    sheet = get_sheet()
    if sheet is None:
        return pd.DataFrame()
    try:
        data = sheet.get_all_records()
        if not data:
            return pd.DataFrame(columns=["Дата", "Товар", "Цена за ед.", "Количество", "Единица", "Сумма", "Поставщик", "Категория", "Примечание"])
        return pd.DataFrame(data)
    except:
        return pd.DataFrame()

def save_data(df):
    sheet = get_sheet()
    if sheet is None:
        return False
    try:
        sheet.clear()
        if not df.empty:
            sheet.update([df.columns.values.tolist()] + df.values.tolist())
        return True
    except Exception as e:
        st.error(f"Ошибка: {e}")
        return False

# === КАТЕГОРИИ ===
CATEGORIES = {
    "Мясо и птица": ["говядина", "куриное филе", "крылышки", "колбаса", "куриные", "марьян"],
    "Молочные продукты": ["сыр", "моцарелла", "сметана", "молоко", "майонез", "сметанковый", "джугас"],
    "Овощи и зелень": ["лук", "чеснок", "помидоры", "огурцы", "морковь", "перец", "картофель", "укроп", "петрушка", "сельдерей", "джусай", "зеленый лук"],
    "Бакалея": ["мука", "масло", "кетчуп", "соус", "лаваш", "яйца", "специи", "приправа", "орегано", "базилик", "лавровый лист", "уксус", "соевый соус", "барбекю", "фритюра", "дрожжи", "томатная паста"],
    "Прочее": []
}

def detect_category(product):
    product = product.lower()
    for cat, keywords in CATEGORIES.items():
        for kw in keywords:
            if kw in product:
                return cat
    return "Прочее"

# === ПРОСТОЙ ПАРСЕР ТЕКСТА ===
def parse_simple_text(text):
    """Простой парсер для текста с закупками"""
    purchases = []
    lines = text.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Ищем числа (потенциальные суммы)
        numbers = re.findall(r'\b(\d{3,6})\b', line)
        if not numbers:
            continue
        
        total = int(numbers[-1])
        
        # Название товара
        product = re.sub(r'^\d+\s+', '', line)
        product = re.sub(r'\s+\d{3,6}\s*$', '', product)
        product = re.sub(r'\d+\s*(?:тг|₸)', '', product)
        product = re.sub(r'\d+(?:[.,]\d+)?\s*(?:кг|г|л|мл|шт|бутылки|пачку|мешка|упаковок|кусочка|лоток|ведер|пакет|пучка|банки)', '', product)
        product = re.sub(r'\s+', ' ', product).strip()
        
        if product and len(product) > 2 and total > 0:
            purchases.append({
                "Товар": product,
                "Сумма": total
            })
    
    return purchases

def generate_report(df, period_name):
    if df.empty:
        return {}
    total_sum = df["Сумма"].sum()
    total_items = len(df)
    total_qty = df["Количество"].sum()
    unique_products = df["Товар"].nunique()
    by_category = df.groupby("Категория")["Сумма"].sum().reset_index()
    by_category = by_category.sort_values("Сумма", ascending=False)
    top_products = df.groupby("Товар")["Сумма"].sum().sort_values(ascending=False).head(10).reset_index()
    top_products.columns = ["Товар", "Сумма"]
    by_supplier = df.groupby("Поставщик")["Сумма"].sum().reset_index()
    by_supplier = by_supplier.sort_values("Сумма", ascending=False)
    daily = df.groupby(df["Дата_парс"].dt.date)["Сумма"].sum().reset_index()
    daily.columns = ["Дата", "Сумма"]
    total_rev = df["Сумма"].sum()
    abc_data = df.groupby("Товар")["Сумма"].sum().reset_index()
    abc_data["Доля"] = abc_data["Сумма"] / total_rev * 100 if total_rev > 0 else 0
    abc_data.sort_values("Сумма", ascending=False, inplace=True)
    abc_data["Категория"] = abc_data["Доля"].apply(lambda x: "A" if x >= 40 else ("B" if x >= 15 else "C"))
    report = {
        "summary": pd.DataFrame([{
            "Период": period_name,
            "Дата формирования": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Общие расходы": f"{total_sum:,.0f} ₸",
            "Количество закупок": total_items,
            "Общее количество": f"{total_qty:.1f} ед.",
            "Уникальных товаров": unique_products,
        }]),
        "by_category": by_category,
        "by_supplier": by_supplier,
        "top_products": top_products,
        "daily": daily,
        "abc": abc_data
    }
    return report

def send_report_to_sheets(df, period_name):
    report_sheet = get_report_sheet()
    if report_sheet is None:
        return False
    report_data = generate_report(df, period_name)
    if not report_data:
        return False
    try:
        report_sheet.clear()
        report_sheet.update_cell(1, 1, "ОТЧЕТ О ЗАКУПКАХ")
        report_sheet.update_cell(2, 1, f"Период: {period_name}")
        report_sheet.update_cell(3, 1, f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        row = 5
        if not report_data["summary"].empty:
            report_sheet.update_cell(row, 1, "1. СВОДКА")
            row += 1
            headers = report_data["summary"].columns.tolist()
            report_sheet.update([headers] + report_data["summary"].values.tolist(), value_input_option="USER_ENTERED")
            row += len(report_data["summary"]) + 1
        report_sheet.update_cell(row, 1, "2. РАСХОДЫ ПО КАТЕГОРИЯМ")
        row += 1
        if not report_data["by_category"].empty:
            cat_headers = report_data["by_category"].columns.tolist()
            report_sheet.update([cat_headers] + report_data["by_category"].values.tolist(), value_input_option="USER_ENTERED")
            row += len(report_data["by_category"]) + 1
        report_sheet.update_cell(row, 1, "3. ТОП-10 ТОВАРОВ")
        row += 1
        if not report_data["top_products"].empty:
            top_headers = report_data["top_products"].columns.tolist()
            report_sheet.update([top_headers] + report_data["top_products"].values.tolist(), value_input_option="USER_ENTERED")
            row += len(report_data["top_products"]) + 1
        report_sheet.update_cell(row, 1, "4. РАСХОДЫ ПО ПОСТАВЩИКАМ")
        row += 1
        if not report_data["by_supplier"].empty:
            sup_headers = report_data["by_supplier"].columns.tolist()
            report_sheet.update([sup_headers] + report_data["by_supplier"].values.tolist(), value_input_option="USER_ENTERED")
            row += len(report_data["by_supplier"]) + 1
        report_sheet.update_cell(row, 1, "5. ABC-АНАЛИЗ ТОВАРОВ")
        row += 1
        if not report_data["abc"].empty:
            abc_data = report_data["abc"][["Товар", "Сумма", "Доля", "Категория"]].values.tolist()
            abc_headers = ["Товар", "Сумма", "Доля %", "Категория"]
            report_sheet.update([abc_headers] + abc_data, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"Ошибка: {e}")
        return False

st.markdown("""
<style>
    /* АДАПТИВНАЯ ТЕМА - подстраивается под систему */
    
    /* === СВЕТЛАЯ ТЕМА (по умолчанию) === */
    .stApp {
        background: #f8f9fa;
    }
    
    /* Текст */
    h1, h2, h3, h4, p, li, .stMarkdown, label, .stCaption {
        color: #212529;
    }
    
    /* Метрики */
    div[data-testid="stMetric"] {
        background: rgba(0, 0, 0, 0.03);
        border-radius: 20px;
        padding: 20px;
        border: 1px solid rgba(0, 0, 0, 0.1);
        box-shadow: 0 2px 8px rgba(0,0,0,0.02);
    }
    div[data-testid="stMetric"] label {
        color: #212529 !important;
    }
    div[data-testid="stMetric"] div {
        color: #212529 !important;
    }
    
    /* Боковая панель */
    [data-testid="stSidebar"] {
        background: #f0f0f0;
        border-right: 1px solid #ddd;
    }
    [data-testid="stSidebar"] * {
        color: #212529 !important;
    }
    
    /* Поля ввода - СВЕТЛАЯ ТЕМА */
    .stTextInput > div > div > input, 
    .stNumberInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div > select,
    .stDateInput > div > div > input,
    .stTimeInput > div > div > input {
        background-color: #ffffff;
        border-radius: 15px;
        border: 1px solid #ced4da;
        font-size: 16px;
        color: #212529 !important;
    }
    .stTextInput > div > div > input:focus, 
    .stNumberInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: #4a90e2;
        box-shadow: 0 0 0 2px rgba(74,144,226,0.2);
        outline: none;
    }
    .stTextInput label, .stNumberInput label, .stTextArea label, .stSelectbox label {
        color: #212529 !important;
        font-weight: 500;
    }
    
    /* Выпадающие списки */
    .stSelectbox > div > div > div {
        background-color: #ffffff;
    }
    
    /* Мультиселект */
    .stMultiSelect div {
        background-color: #ffffff;
        color: #212529;
    }
    
    /* Кнопки */
    .stButton > button {
        background: #4a90e2;
        color: white;
        border: none;
        border-radius: 30px;
        padding: 10px 24px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        transform: scale(1.02);
        background: #5a9ee2;
        color: white;
    }
    
    /* Таблицы */
    .stDataFrame {
        background: #ffffff;
        border-radius: 15px;
        border: 1px solid #dee2e6;
    }
    .stDataFrame th {
        background: #4a90e2 !important;
        color: white !important;
        font-weight: bold;
    }
    .stDataFrame td {
        color: #212529 !important;
    }
    
    /* Вкладки */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: rgba(0, 0, 0, 0.04);
        border-radius: 30px;
        padding: 5px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 25px;
        padding: 8px 20px;
        font-weight: bold;
        color: #212529;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4a90e2 !important;
        color: white !important;
    }
    
    /* Информационные блоки */
    .stInfo {
        background-color: #e9ecef !important;
        color: #212529 !important;
        border-left: 4px solid #4a90e2;
    }
    
    /* Слайдеры */
    .stSlider > div > div > div {
        background-color: #4a90e2;
    }
    
    /* Чекбоксы и радиокнопки */
    .stCheckbox label, .stRadio label {
        color: #212529 !important;
    }
    
    /* === ТЁМНАЯ ТЕМА (автоматически) === */
    @media (prefers-color-scheme: dark) {
        .stApp {
            background: #1a1a1a;
        }
        
        h1, h2, h3, h4, p, li, .stMarkdown, label, .stCaption {
            color: #e0e0e0;
        }
        
        div[data-testid="stMetric"] {
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid rgba(255, 255, 255, 0.15);
        }
        div[data-testid="stMetric"] label {
            color: #e0e0e0 !important;
        }
        div[data-testid="stMetric"] div {
            color: #e0e0e0 !important;
        }
        
        [data-testid="stSidebar"] {
            background: #2a2a2a;
            border-right: 1px solid #444;
        }
        [data-testid="stSidebar"] * {
            color: #e0e0e0 !important;
        }
        
        /* Поля ввода - ТЁМНАЯ ТЕМА */
        .stTextInput > div > div > input, 
        .stNumberInput > div > div > input,
        .stTextArea > div > div > textarea,
        .stSelectbox > div > div > select,
        .stDateInput > div > div > input,
        .stTimeInput > div > div > input {
            background-color: #2d2d2d;
            border-color: #555;
            color: #e0e0e0 !important;
        }
        .stTextInput > div > div > input:focus, 
        .stNumberInput > div > div > input:focus,
        .stTextArea > div > div > textarea:focus {
            border-color: #4a90e2;
        }
        .stTextInput label, .stNumberInput label, .stTextArea label, .stSelectbox label {
            color: #e0e0e0 !important;
        }
        
        .stSelectbox > div > div > div {
            background-color: #2d2d2d;
            color: #e0e0e0;
        }
        
        .stMultiSelect div {
            background-color: #2d2d2d;
            color: #e0e0e0;
        }
        
        .stDataFrame {
            background: #2d2d2d;
            border-color: #444;
        }
        .stDataFrame td {
            color: #e0e0e0 !important;
        }
        
        .stTabs [data-baseweb="tab-list"] {
            background-color: rgba(255, 255, 255, 0.05);
        }
        .stTabs [data-baseweb="tab"] {
            color: #e0e0e0;
        }
        
        .stInfo {
            background-color: #2a2a2a !important;
            color: #e0e0e0 !important;
        }
        
        .stCheckbox label, .stRadio label {
            color: #e0e0e0 !important;
        }
    }
    
    /* === ОБЩИЕ СТИЛИ (работают в обеих темах) === */
    .stButton > button {
        background: #4a90e2;
        color: white;
    }
    .stButton > button:hover {
        background: #5a9ee2;
    }
    
    .stDataFrame th {
        background: #4a90e2 !important;
        color: white !important;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #4a90e2 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📦 Учет закупок кафе")
st.caption("Данные сохраняются в Google Sheets")

df = load_data()
if df.empty:
    df = pd.DataFrame(columns=["Дата", "Товар", "Цена за ед.", "Количество", "Единица", "Сумма", "Поставщик", "Категория", "Примечание"])
    save_data(df)

# === БОКОВАЯ ПАНЕЛЬ ===
with st.sidebar:
    st.markdown("### 📅 Период")
    period = st.radio("Показать", ["Сегодня", "Неделя", "Месяц", "Всё время"])
    st.markdown("---")
    
    # === ЗАГРУЗКА EXCEL ===
    st.markdown("### 📎 Загрузить Excel")
    uploaded_file = st.file_uploader("Файл .xlsx или .xls", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df_upload = pd.read_excel(uploaded_file)
            st.success(f"Загружено: {len(df_upload)} строк")
            with st.expander("Просмотр данных"):
                st.dataframe(df_upload.head(10), use_container_width=True)
            if st.button("📊 Добавить данные из Excel"):
                added = 0
                for idx, row in df_upload.iterrows():
                    try:
                        product = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
                        if not product or product.lower() in ['товар', 'название', 'nan']:
                            continue
                        price = 0
                        try:
                            price = float(str(row.iloc[1]).replace(',', '.').replace('₸', '').replace('тг', '').strip())
                        except:
                            price = 0
                        qty = 1
                        try:
                            qty = float(str(row.iloc[2]).replace(',', '.').strip())
                        except:
                            qty = 1
                        total = 0
                        try:
                            total = float(str(row.iloc[3]).replace(',', '.').replace('₸', '').replace('тг', '').strip())
                        except:
                            total = 0
                        if total == 0 and price > 0 and qty > 0:
                            total = price * qty
                        if total > 0:
                            new_row = pd.DataFrame([{
                                "Дата": datetime.now().strftime("%d.%m.%Y"),
                                "Товар": product[:100],
                                "Цена за ед.": round(price, 2),
                                "Количество": qty,
                                "Единица": "шт",
                                "Сумма": round(total, 2),
                                "Поставщик": "Замира",
                                "Категория": detect_category(product),
                                "Примечание": ""
                            }])
                            df = pd.concat([df, new_row], ignore_index=True)
                            added += 1
                    except:
                        pass
                if added > 0:
                    save_data(df)
                    st.success(f"✅ Добавлено {added} позиций!")
                    st.rerun()
                else:
                    st.error("Не удалось добавить позиции")
        except Exception as e:
            st.error(f"Ошибка: {e}")
    
    st.markdown("---")
    
    # === ВСТАВКА ТЕКСТА ===
    st.markdown("### 📝 Вставить текст из документа")
    st.info("Скопируйте текст из Word-файла и вставьте сюда")
    
    manual_text = st.text_area("Вставьте текст с закупками:", height=200)
    
    if st.button("📊 Обработать текст", key="parse_manual"):
        if manual_text:
            purchases = parse_simple_text(manual_text)
            if purchases:
                added = 0
                for p in purchases:
                    new_row = pd.DataFrame([{
                        "Дата": datetime.now().strftime("%d.%m.%Y"),
                        "Товар": p["Товар"],
                        "Цена за ед.": p.get("Цена за ед.", 0),
                        "Количество": p.get("Количество", 1),
                        "Единица": p.get("Единица", "шт"),
                        "Сумма": p["Сумма"],
                        "Поставщик": "Замира",
                        "Категория": detect_category(p["Товар"]),
                        "Примечание": ""
                    }])
                    df = pd.concat([df, new_row], ignore_index=True)
                    added += 1
                if added > 0:
                    save_data(df)
                    st.success(f"✅ Добавлено {added} позиций!")
                    st.rerun()
                else:
                    st.error("Не удалось распознать данные")
            else:
                st.error("Не найдено данных в тексте")
    
    st.markdown("---")
    
    # === РУЧНОЕ ДОБАВЛЕНИЕ ===
    st.markdown("### ➕ Ручное добавление")
    with st.form("add_purchase"):
        product = st.text_input("Товар")
        col1, col2 = st.columns(2)
        with col1:
            price = st.number_input("Цена за ед.", min_value=0.0, value=0.0, step=10.0)
            qty = st.number_input("Количество", min_value=0.0, value=1.0, step=0.1)
            unit = st.selectbox("Единица", ["кг", "г", "л", "мл", "шт"])
        with col2:
            total = st.number_input("Сумма", min_value=0.0, value=0.0, step=100.0)
            supplier = st.selectbox("Поставщик", ["Замира", "Магнум", "Метро", "Другое"])
            note = st.text_area("Примечание")
        if st.form_submit_button("Добавить"):
            if product:
                if total == 0 and price > 0 and qty > 0:
                    total = price * qty
                category = detect_category(product)
                new_row = pd.DataFrame([{
                    "Дата": datetime.now().strftime("%d.%m.%Y"),
                    "Товар": product,
                    "Цена за ед.": round(price, 2),
                    "Количество": qty,
                    "Единица": unit,
                    "Сумма": round(total, 2),
                    "Поставщик": supplier,
                    "Категория": category,
                    "Примечание": note
                }])
                df = pd.concat([df, new_row], ignore_index=True)
                save_data(df)
                st.success("✅ Добавлено!")
                st.rerun()

# === ФИЛЬТРАЦИЯ ===
if not df.empty and "Дата" in df.columns:
    df["Дата_парс"] = pd.to_datetime(df["Дата"], format="%d.%m.%Y", errors="coerce")
    today = datetime.now().date()
    if period == "Сегодня":
        filtered = df[df["Дата_парс"].dt.date == today]
    elif period == "Неделя":
        filtered = df[df["Дата_парс"].dt.date >= today - timedelta(days=7)]
    elif period == "Месяц":
        filtered = df[df["Дата_парс"].dt.date >= today - timedelta(days=30)]
    else:
        filtered = df.copy()
else:
    filtered = df

# === МЕТРИКИ ===
c1, c2, c3, c4 = st.columns(4)
if not filtered.empty:
    c1.metric("💰 Расходы", f"{filtered['Сумма'].sum():,.0f} ₸")
    c2.metric("📦 Закупок", len(filtered))
    c3.metric("📊 Количество", f"{filtered['Количество'].sum():,.1f} ед.")
    c4.metric("🏷️ Товаров", filtered['Товар'].nunique())
else:
    c1.metric("💰 Расходы", "0 ₸")
    c2.metric("📦 Закупок", "0")
    c3.metric("📊 Количество", "0")
    c4.metric("🏷️ Товаров", "0")

st.markdown("---")

# === КНОПКИ ОТЧЁТОВ ===
if not filtered.empty:
    col_r1, col_r2, col_r3 = st.columns(3)
    with col_r1:
        if st.button("📊 Сформировать отчёт", use_container_width=True):
            report_data = generate_report(filtered, period)
            st.session_state.report = report_data
            st.success("✅ Отчёт сформирован!")
    with col_r2:
        if st.button("📤 Выгрузить отчёт в Sheets", use_container_width=True):
            if send_report_to_sheets(filtered, period):
                st.success("✅ Отчёт выгружен в Google Sheets (лист 'ОТЧЕТЫ')!")
            else:
                st.error("Ошибка выгрузки")
    with col_r3:
        report_data = generate_report(filtered, period)
        if report_data and not report_data["summary"].empty:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                report_data["summary"].to_excel(writer, sheet_name="Сводка", index=False)
                report_data["by_category"].to_excel(writer, sheet_name="По категориям", index=False)
                report_data["top_products"].to_excel(writer, sheet_name="Топ товаров", index=False)
                report_data["by_supplier"].to_excel(writer, sheet_name="По поставщикам", index=False)
                report_data["daily"].to_excel(writer, sheet_name="По дням", index=False)
                report_data["abc"].to_excel(writer, sheet_name="ABC анализ", index=False)
            st.download_button("📥 Скачать отчёт (Excel)", output.getvalue(), f"отчет_{datetime.now().strftime('%Y-%m-%d')}.xlsx", use_container_width=True)

st.markdown("---")

# === ВКЛАДКИ ===
tab1, tab2, tab3, tab4 = st.tabs(["📊 По дням", "🍽️ По категориям", "📦 По товарам", "📈 Отчёты"])

with tab1:
    if not filtered.empty and "Дата_парс" in filtered.columns:
        daily = filtered.groupby(filtered["Дата_парс"].dt.date)["Сумма"].sum().reset_index()
        daily.columns = ["Дата", "Сумма"]
        if not daily.empty:
            fig = px.line(daily, x="Дата", y="Сумма", template="plotly_dark", color_discrete_sequence=["#d4a373"])
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

with tab2:
    if not filtered.empty:
        by_cat = filtered.groupby("Категория")["Сумма"].sum().reset_index()
        fig = px.pie(by_cat, values="Сумма", names="Категория", template="plotly_dark")
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(by_cat.sort_values("Сумма", ascending=False), use_container_width=True)

with tab3:
    if not filtered.empty:
        by_product = filtered.groupby("Товар")["Сумма"].sum().sort_values(ascending=False).head(15).reset_index()
        fig = px.bar(by_product, x="Сумма", y="Товар", orientation='h', template="plotly_dark", color_discrete_sequence=["#d4a373"])
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=500)
        st.plotly_chart(fig, use_container_width=True)

with tab4:
    if "report" in st.session_state and st.session_state.report:
        r = st.session_state.report
        st.subheader("📊 Сводка")
        st.dataframe(r["summary"], use_container_width=True)
        st.subheader("📈 Расходы по категориям")
        st.dataframe(r["by_category"], use_container_width=True)
        st.subheader("🏆 Топ-10 товаров")
        st.dataframe(r["top_products"], use_container_width=True)
        st.subheader("🏪 Расходы по поставщикам")
        st.dataframe(r["by_supplier"], use_container_width=True)
        st.subheader("📋 ABC-анализ")
        st.dataframe(r["abc"][["Товар", "Сумма", "Доля", "Категория"]], use_container_width=True)
        abc_chart = r["abc"].groupby("Категория")["Сумма"].sum().reset_index()
        if not abc_chart.empty:
            fig = px.pie(abc_chart, values="Сумма", names="Категория", template="plotly_dark", title="ABC-анализ")
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("📭 Нажмите 'Сформировать отчёт' для просмотра аналитики")

st.caption("💾 Данные в Google Sheets | Отчёты можно выгрузить в Excel или Google Sheets")
