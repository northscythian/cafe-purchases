import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials

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

CATEGORIES = {
    "Мясо и птица": ["говядина", "куриное филе", "крылышки", "колбаса"],
    "Молочные продукты": ["сыр", "моцарелла", "сметана", "молоко", "майонез"],
    "Овощи и зелень": ["лук", "чеснок", "помидоры", "огурцы", "морковь", "перец", "картофель"],
    "Бакалея": ["мука", "масло", "кетчуп", "соус", "лаваш", "яйца", "специи"],
    "Прочее": []
}

def detect_category(product):
    product = product.lower()
    for cat, keywords in CATEGORIES.items():
        for kw in keywords:
            if kw in product:
                return cat
    return "Прочее"

def get_default_data():
    default_data = [
        (datetime.now().strftime("%d.%m.%Y"), "Масло для фритюра", 1042.86, 21.0, "л", 21900, "Замира", "Бакалея", ""),
        (datetime.now().strftime("%d.%m.%Y"), "Дрожжи", 4.4, 50.0, "г", 220, "Замира", "Бакалея", ""),
        (datetime.now().strftime("%d.%m.%Y"), "Масло растительное", 760.0, 20.0, "л", 15200, "Замира", "Бакалея", ""),
    ]
    return pd.DataFrame(default_data, columns=["Дата", "Товар", "Цена за ед.", "Количество", "Единица", "Сумма", "Поставщик", "Категория", "Примечание"])

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
    with tab3:
        st.dataframe(filtered, use_container_width=True)
else:
    st.info("Нет данных")

st.caption("💾 Данные в Google Sheets")
