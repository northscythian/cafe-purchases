import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

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

# === РАСШИРЕННЫЕ КАТЕГОРИИ ===
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

# === ПАРСИНГ ЦЕНЫ ЗА ЕДИНИЦУ ===
def parse_price_per_unit(price_str):
    price_str = str(price_str).lower().strip()
    main_price = re.search(r'(\d+(?:[.,]\d+)?)\s*тг', price_str)
    if not main_price:
        main_price = re.search(r'(\d+(?:[.,]\d+)?)\s*₸', price_str)
    if not main_price:
        return 0, "шт"
    
    price = float(main_price.group(1).replace(',', '.'))
    
    per_kg = re.search(r'за\s+(\d+(?:[.,]\d+)?)\s*кг', price_str)
    if per_kg:
        qty = float(per_kg.group(1).replace(',', '.'))
        return round(price / qty, 2), "кг"
    
    per_l = re.search(r'за\s+(\d+(?:[.,]\d+)?)\s*л', price_str)
    if per_l:
        qty = float(per_l.group(1).replace(',', '.'))
        return round(price / qty, 2), "л"
    
    if "/кг" in price_str or "тг/кг" in price_str:
        return price, "кг"
    elif "/л" in price_str or "тг/л" in price_str:
        return price, "л"
    
    return price, "шт"

# === ПАРСИНГ КОЛИЧЕСТВА ===
def parse_quantity(qty_str):
    qty_str = str(qty_str).lower().strip()
    
    kg_match = re.search(r'(\d+(?:[.,]\d+)?)\s*кг', qty_str)
    g_match = re.search(r'(\d+(?:[.,]\d+)?)\s*г', qty_str)
    
    if kg_match or g_match:
        total = 0
        if kg_match:
            total += float(kg_match.group(1).replace(',', '.'))
        if g_match:
            total += float(g_match.group(1).replace(',', '.')) / 1000
        return round(total, 3), "кг"
    
    l_match = re.search(r'(\d+(?:[.,]\d+)?)\s*л', qty_str)
    if l_match:
        return float(l_match.group(1).replace(',', '.')), "л"
    
    pcs_match = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:шт|штук|пачк|упаковк|бутылк|банк|мешк|кусочк|лоток|ведер|пакет|пучк)', qty_str)
    if pcs_match:
        return float(pcs_match.group(1).replace(',', '.')), "шт"
    
    num_match = re.search(r'(\d+(?:[.,]\d+)?)', qty_str)
    if num_match:
        return float(num_match.group(1).replace(',', '.')), "шт"
    
    return 1, "шт"

# === ВСЕ 59 ПОЗИЦИЙ (РАССЧИТАНЫ АВТОМАТИЧЕСКИ) ===
def get_all_purchases():
    raw_data = [
        ("Масло для фритюра", "7300 тг за 7 л", "3 бутылки", 21900),
        ("Дрожжи", "220 тг за пачку 50 г", "1 пачка", 220),
        ("Масло растительное", "3800 тг за 5 л", "4 бутылки", 15200),
        ("Сыр сметанковый", "2500 тг/кг", "4 кг 300 г", 10750),
        ("Мука", "7950 тг за 25 кг", "2 мешка", 15900),
        ("Сметана Умит", "440 тг за 500 г", "5 упаковок", 2200),
        ("Сыр Джугас 40%", "3500 тг за 180 г", "2 кусочка", 7000),
        ("Моцарелла", "3400 тг/кг", "2 кг", 6800),
        ("Моцарелла (другой бренд)", "3000 тг/кг", "4 кг", 12000),
        ("Крылышки куриные", "1550 тг/кг", "15 кг 915 г", 24668),
        ("Куриное филе", "1500 тг/кг", "14 кг", 21000),
        ("Колбаса Марьям", "1850 тг/кг", "2 кг 180 г", 4033),
        ("Соус Барбекю", "1600 тг за 0,5 л", "1 упаковка", 1600),
        ("Масло растительное", "3700 тг за 5 л", "2 бутылки", 7400),
        ("Томатная паста", "880 тг за 1 л", "3 банки", 2640),
        ("Колбаса Марьям", "1300 тг/кг", "3 кг 940 г", 5122),
        ("Сыр сметанковый", "2500 тг/кг", "5 кг 860 г", 14650),
        ("Моцарелла", "2600 тг/кг", "4 кг", 10400),
        ("Сметана", "2000 тг за 2 л", "2 л", 2000),
        ("Лук", "230 тг/кг", "8 кг 430 г", 1940),
        ("Чеснок", "800 тг/кг", "1 кг 14 г", 811),
        ("Чесночные дудки", "800 тг/кг", "610 г", 485),
        ("Сельдерей", "180 тг за пучок", "2 пучка", 360),
        ("Помидоры", "550 тг/кг", "2 кг 590 г", 1450),
        ("Говядина", "3800 тг/кг", "1 кг 800 г", 6850),
        ("Майонез", "1150 тг за 1 л", "5 ведер", 5750),
        ("Кетчуп", "950 тг за 1 л", "2 ведра", 1900),
        ("Кетчуп тетрапак", "715 тг за 0,5 л", "5 шт", 3575),
        ("Лаваш", "650 тг за пакет", "4 пакета", 2600),
        ("Помидоры", "550 тг/кг", "2,5 кг", 1375),
        ("Огурцы", "550 тг/кг", "2,5 кг", 1375),
        ("Сельдерей", "150 тг за пучок", "2 пучка", 300),
        ("Морковь", "350 тг/кг", "2 кг 700 г", 945),
        ("Лук", "200 тг/кг", "38 кг 100 г", 7620),
        ("Свежий перец", "1100 тг/кг", "2 кг 95 г", 2305),
        ("Морковь", "380 тг/кг", "2 кг", 760),
        ("Джусай", "175 тг за пучок", "1 пучок", 175),
        ("Зеленый лук", "150 тг за пучок", "1 пучок", 150),
        ("Укроп", "100 тг за пучок", "1 пучок", 100),
        ("Картофель", "320 тг/кг", "8 кг 775 г", 2810),
        ("Помидоры", "780 тг/кг", "2 кг 64 г", 1609),
        ("Огурцы", "400 тг/кг", "2 кг 400 г", 960),
        ("Яйца", "1650 тг за лоток", "30 шт", 1650),
        ("Молоко", "450 тг/л", "2 л", 900),
        ("Томатная паста", "800 тг за 1 л", "2 банки", 1600),
        ("Говядина", "3793 тг/кг", "5 кг 730 г", 21735),
        ("Говядина", "3677 тг/кг", "5 кг 44 г", 18549),
        ("Картофель фри", "2173 тг/кг", "8 кг", 17384),
        ("Черный перец", "5586 тг/кг", "1 кг", 5586),
        ("Чеснок сушеный", "5027 тг/кг", "2 кг", 10055),
        ("Красный перец", "1787 тг/кг", "2 кг", 3575),
        ("Орегано", "6704 тг/кг", "1 кг", 6704),
        ("Базилик", "5586 тг/кг", "1 кг", 5586),
        ("Универсальная приправа", "5028 тг/кг", "1 кг", 5028),
        ("Китайский уксус", "3966 тг за 3 л", "1 бутылка", 3966),
        ("Соевый соус", "1117 тг за 0,5 л", "2 бутылки", 2234),
        ("Лавровый лист", "500 тг за упаковку", "1 упаковка", 500),
        ("Помидоры", "600 тг/кг", "2 кг 430 г", 1458),
        ("Огурцы", "по записи сумма 1311 тг", "2 кг 185 г", 1311),
    ]
    
    purchases = []
    for product, price_str, qty_str, total in raw_data:
        price_per_unit, _ = parse_price_per_unit(price_str)
        quantity, unit = parse_quantity(qty_str)
        purchases.append({
            "Товар": product,
            "Цена за ед.": price_per_unit,
            "Количество": quantity,
            "Единица": unit,
            "Сумма": total,
            "Категория": detect_category(product),
        })
    return purchases

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
st.caption("Данные сохраняются в Google Sheets")

df = load_data()
if df.empty:
    df = get_default_data()
    save_data(df)

# === КНОПКА ЗАГРУЗКИ ВСЕХ 59 ПОЗИЦИЙ ===
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("📋 ЗАГРУЗИТЬ ВСЕ 59 ПОЗИЦИЙ", use_container_width=True):
        purchases = get_all_purchases()
        for p in purchases:
            p["Дата"] = datetime.now().strftime("%d.%m.%Y")
            p["Поставщик"] = "Замира"
            p["Примечание"] = ""
        
        new_df = pd.DataFrame(purchases)
        cols_order = ["Дата", "Товар", "Категория", "Цена за ед.", "Количество", "Единица", "Сумма", "Поставщик", "Примечание"]
        new_df = new_df[cols_order]
        df = pd.concat([df, new_df], ignore_index=True)
        save_data(df)
        st.success(f"✅ Загружено {len(purchases)} позиций!")
        st.rerun()

with st.sidebar:
    st.markdown("### 📅 Период")
    period = st.radio("Показать", ["Сегодня", "Неделя", "Месяц", "Всё время"])
    
    st.markdown("---")
    st.markdown("### ➕ Добавить закупку")
    
    with st.form("add_purchase"):
        product = st.text_input("Товар")
        col1, col2, col3 = st.columns(3)
        with col1:
            price = st.number_input("Цена за ед.", min_value=0.0, value=0.0, step=10.0)
            qty = st.number_input("Количество", min_value=0.0, value=1.0, step=0.1)
        with col2:
            unit = st.selectbox("Единица", ["кг", "г", "л", "мл", "шт"])
            total = st.number_input("Сумма", min_value=0.0, value=0.0, step=100.0)
        with col3:
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
                if save_data(df):
                    st.success("Добавлено!")
                    st.rerun()

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

if not filtered.empty:
    tab1, tab2, tab3 = st.tabs(["📊 По дням", "🍽️ По категориям", "📋 Список"])
    with tab1:
        if "Дата_парс" in filtered.columns:
            daily = filtered.groupby(filtered["Дата_парс"].dt.date)["Сумма"].sum().reset_index()
            daily.columns = ["Дата", "Сумма"]
            fig = px.line(daily, x="Дата", y="Сумма", template="plotly_dark", color_discrete_sequence=["#d4a373"])
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
    with tab2:
        by_cat = filtered.groupby("Категория")["Сумма"].sum().reset_index()
        fig = px.pie(by_cat, values="Сумма", names="Категория", template="plotly_dark")
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(by_cat.sort_values("Сумма", ascending=False), use_container_width=True)
    with tab3:
        st.dataframe(filtered, use_container_width=True)
else:
    st.info("Нет данных. Нажмите 'ЗАГРУЗИТЬ ВСЕ 59 ПОЗИЦИЙ' или добавьте вручную.")

st.caption("💾 Данные в Google Sheets")
