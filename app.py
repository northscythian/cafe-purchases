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
    "Хозтовары": ["перчатки", "тряпки", "губки", "салфетки", "кнопки", "маркер", "тетрадь", "мелки", "посуда", "сковорода", "тарелки", "миски", "поварешка", "калькулятор"],
    "Аптечка": ["бинт", "перекись", "спирт", "зеленка", "марля", "ватные палочки"],
    "Прочее": []
}

def detect_category(product):
    product = product.lower()
    for cat, keywords in CATEGORIES.items():
        for kw in keywords:
            if kw in product:
                return cat
    return "Прочее"

def detect_unit(product):
    """Определяет единицу измерения по названию товара"""
    product = product.lower()
    if any(x in product for x in ["перчатки", "тряпки", "салфетки", "кнопки", "маркер", "тетрадь", "мелки", "тарелки", "миски", "поварешка", "калькулятор", "рулон"]):
        return "шт"
    if any(x in product for x in ["бинт", "марля"]):
        return "м"
    if any(x in product for x in ["спирт", "перекись", "зеленка"]):
        return "мл"
    return "шт"

def smart_parse_excel(df_upload):
    """Умный парсер Excel файла — определяет структуру и корректно заполняет данные"""
    required_cols = ["Дата", "Товар", "Цена за ед.", "Количество", "Единица", "Сумма", "Поставщик", "Категория", "Примечаение"]
    
    # Словарь соответствия колонок (поддерживает разные названия)
    col_mapping = {
        "Товар": ["товар", "название", "продукт", "наименование", "item", "product"],
        "Цена за ед.": ["цена за ед", "цена за ед.", "цена", "price", "цена за кг", "цена за шт"],
        "Количество": ["количество", "кол-во", "qty", "quantity", "кол"],
        "Единица": ["единица", "ед", "ед.", "unit"],
        "Сумма": ["сумма", "total", "сумма ₸", "сумма (₸)"],
        "Поставщик": ["поставщик", "supplier", "поставщик"],
        "Категория": ["категория", "category", "кат"],
        "Дата": ["дата", "date", "день"],
        "Примечание": ["примечание", "note", "комментарий"]
    }
    
    # Определяем, какие колонки есть в файле
    detected = {}
    for target, variants in col_mapping.items():
        for col in df_upload.columns:
            col_lower = str(col).lower().strip()
            for variant in variants:
                if variant in col_lower:
                    detected[target] = col
                    break
            if target in detected:
                break
    
    # Если не нашли "Товар" — предположим, что это первая колонка
    if "Товар" not in detected and len(df_upload.columns) > 0:
        detected["Товар"] = df_upload.columns[0]
    
    result = []
    for idx, row in df_upload.iterrows():
        try:
            # Получаем товар
            product = str(row[detected["Товар"]]) if "Товар" in detected and pd.notna(row[detected["Товар"]]) else ""
            if not product or product.lower() in ['товар', 'название', 'nan', '']:
                continue
            
            # Получаем цену
            price = 0
            if "Цена за ед." in detected and pd.notna(row[detected["Цена за ед."]]):
                try:
                    price = float(str(row[detected["Цена за ед."]]).replace(',', '.').replace('₸', '').replace('тг', '').strip())
                except:
                    price = 0
            
            # Получаем количество
            qty = 1
            if "Количество" in detected and pd.notna(row[detected["Количество"]]):
                try:
                    qty = float(str(row[detected["Количество"]]).replace(',', '.').strip())
                except:
                    qty = 1
            
            # Получаем единицу измерения
            unit = "шт"
            if "Единица" in detected and pd.notna(row[detected["Единица"]]):
                unit = str(row[detected["Единица"]]).strip()
            else:
                unit = detect_unit(product)
            
            # Получаем сумму
            total = 0
            if "Сумма" in detected and pd.notna(row[detected["Сумма"]]):
                try:
                    total = float(str(row[detected["Сумма"]]).replace(',', '.').replace('₸', '').replace('тг', '').strip())
                except:
                    total = 0
            
            # Если сумма не указана, но есть цена и количество
            if total == 0 and price > 0 and qty > 0:
                total = price * qty
            # Если цена не указана, но есть сумма и количество
            if price == 0 and total > 0 and qty > 0:
                price = total / qty
            
            # Получаем поставщика
            supplier = "Замира"
            if "Поставщик" in detected and pd.notna(row[detected["Поставщик"]]):
                supplier = str(row[detected["Поставщик"]]).strip()
            
            # Получаем категорию (если есть)
            category = None
            if "Категория" in detected and pd.notna(row[detected["Категория"]]):
                category = str(row[detected["Категория"]]).strip()
            if not category:
                category = detect_category(product)
            
            # Получаем дату
            date = datetime.now().strftime("%d.%m.%Y")
            if "Дата" in detected and pd.notna(row[detected["Дата"]]):
                try:
                    d = pd.to_datetime(row[detected["Дата"]], errors='coerce')
                    if pd.notna(d):
                        date = d.strftime("%d.%m.%Y")
                except:
                    pass
            
            if total > 0:
                result.append({
                    "Дата": date,
                    "Товар": product[:100],
                    "Цена за ед.": round(price, 2),
                    "Количество": qty,
                    "Единица": unit,
                    "Сумма": round(total, 2),
                    "Поставщик": supplier,
                    "Категория": category,
                    "Примечание": ""
                })
        except Exception as e:
            continue
    
    return pd.DataFrame(result)

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
    .stApp { background: linear-gradient(135deg, #2d2b2a 0%, #1a1a1a 100%); }
    div[data-testid="stMetric"] { background: linear-gradient(135deg, #3d3a38, #2c2a28); border-radius: 20px; padding: 20px; border: 1px solid #d4a373; }
    h1, h2, h3, h4 { color: #d4a373 !important; }
    .stButton > button { background: linear-gradient(135deg, #d4a373, #b5835a); color: white; border-radius: 30px; }
    .stDataFrame { background: #2a2a2a; border-radius: 15px; border: 1px solid #d4a373; }
    .stDataFrame th { background: #d4a373 !important; color: #1a1a1a !important; }
    .stDataFrame td { color: white !important; }
</style>
""", unsafe_allow_html=True)

st.title("📦 Учет закупок кафе")
st.caption("Данные сохраняются в Google Sheets | Поддерживает любые форматы Excel")

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
    st.info("Поддерживаются любые форматы: приложение само определит колонки")
    uploaded_file = st.file_uploader("Файл .xlsx или .xls", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df_upload = pd.read_excel(uploaded_file)
            st.success(f"Загружено: {len(df_upload)} строк")
            with st.expander("Просмотр загруженных данных"):
                st.dataframe(df_upload.head(10), use_container_width=True)
            if st.button("📊 Умная обработка и добавление", use_container_width=True):
                processed_df = smart_parse_excel(df_upload)
                if not processed_df.empty:
                    df = pd.concat([df, processed_df], ignore_index=True)
                    save_data(df)
                    st.success(f"✅ Добавлено {len(processed_df)} позиций!")
                    st.rerun()
                else:
                    st.error("Не удалось распознать данные. Проверьте формат файла.")
        except Exception as e:
            st.error(f"Ошибка: {e}")
    
    st.markdown("---")
    
    # === РУЧНОЕ ДОБАВЛЕНИЕ ===
    st.markdown("### ➕ Ручное добавление")
    with st.form("add_purchase"):
        product = st.text_input("Товар")
        col1, col2 = st.columns(2)
        with col1:
            price = st.number_input("Цена за ед.", min_value=0.0, value=0.0, step=10.0)
            qty = st.number_input("Количество", min_value=0.0, value=1.0, step=0.1)
            unit = st.selectbox("Единица", ["кг", "г", "л", "мл", "шт", "м"])
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
