import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import re
import os
from io import BytesIO

st.set_page_config(page_title="Учет закупок кафе", page_icon="📦", layout="wide")

DATA_FILE = "покупки_кафе.xlsx"

CATEGORIES = {
    "Мясо и птица": ["говядина", "куриное филе", "крылышки", "колбаса", "куриные", "марьян"],
    "Молочные продукты": ["сыр", "моцарелла", "сметана", "молоко", "майонез", "сметанковый", "джугас", "мопанелла"],
    "Овощи и зелень": ["лук", "чеснок", "помидоры", "огурцы", "морковь", "перец", "картофель", "укроп", "петрушка", "сельдерей", "джусай", "зеленый лук", "чесночные дудки"],
    "Бакалея и масла": ["мука", "масло", "кетчуп", "томатная паста", "соус", "лаваш", "яйца", "специи", "приправа", "орегано", "базилик", "перец", "чеснок сушеный", "лавровый лист", "уксус", "соевый соус", "барбекю", "фритюра", "дрожжи"],
    "Замороженные продукты": ["картофель фри"]
}

def detect_category(product_name):
    product_name = product_name.lower()
    for category, keywords in CATEGORIES.items():
        for keyword in keywords:
            if keyword in product_name:
                return category
    return "Прочее"

def load_purchases():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE)
        # Убеждаемся, что колонки правильные
        required_cols = ["Дата", "Товар", "Категория", "Цена за ед. (₸)", "Количество", "Единица", "Сумма (₸)", "Поставщик", "Примечание"]
        for col in required_cols:
            if col not in df.columns:
                df[col] = None
        return df
    else:
        return pd.DataFrame(columns=[
            "Дата", "Товар", "Категория", "Цена за ед. (₸)", 
            "Количество", "Единица", "Сумма (₸)", "Поставщик", "Примечание"
        ])

def save_purchases(df):
    df.to_excel(DATA_FILE, index=False)

def parse_quantity_to_kg(qty_str):
    """Преобразует количество в кг (для удобства анализа)"""
    qty_str = str(qty_str).lower().strip()
    total_kg = 0
    
    # Поиск кг
    kg_match = re.search(r'(\d+(?:[.,]\d+)?)\s*кг', qty_str)
    if kg_match:
        total_kg += float(kg_match.group(1).replace(',', '.'))
    
    # Поиск г
    g_match = re.search(r'(\d+(?:[.,]\d+)?)\s*г', qty_str)
    if g_match:
        total_kg += float(g_match.group(1).replace(',', '.')) / 1000
    
    # Поиск л (для жидкостей)
    l_match = re.search(r'(\d+(?:[.,]\d+)?)\s*л', qty_str)
    if l_match:
        return float(l_match.group(1).replace(',', '.')), "л"
    
    # Поиск мл
    ml_match = re.search(r'(\d+(?:[.,]\d+)?)\s*мл', qty_str)
    if ml_match:
        return float(ml_match.group(1).replace(',', '.')) / 1000, "л"
    
    # Поиск шт
    pcs_match = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:шт|штук|пачк|упаковк|бутылк|банк|мешк|кусочк|лоток|ведер|пакет|пучк)', qty_str)
    if pcs_match:
        return float(pcs_match.group(1).replace(',', '.')), "шт"
    
    # Просто число
    num_match = re.search(r'(\d+(?:[.,]\d+)?)', qty_str)
    if num_match:
        return float(num_match.group(1).replace(',', '.')), "шт"
    
    return 1, "шт"

def parse_price_per_unit(price_str):
    """Извлекает цену за единицу (кг/л/шт)"""
    price_str = str(price_str).lower().strip()
    
    # Ищем основную сумму
    main_price = re.search(r'(\d+(?:[.,]\d+)?)\s*тг', price_str)
    if not main_price:
        main_price = re.search(r'(\d+(?:[.,]\d+)?)\s*₸', price_str)
    
    if not main_price:
        return 0, "шт"
    
    price = float(main_price.group(1).replace(',', '.'))
    
    # Проверяем, за какое количество эта цена
    # Цена за кг
    per_kg = re.search(r'за\s+(\d+(?:[.,]\d+)?)\s*кг', price_str)
    if per_kg:
        qty = float(per_kg.group(1).replace(',', '.'))
        return round(price / qty, 2), "кг"
    
    # Цена за г
    per_g = re.search(r'за\s+(\d+(?:[.,]\d+)?)\s*г', price_str)
    if per_g:
        qty = float(per_g.group(1).replace(',', '.'))
        return round(price / qty, 2), "г"
    
    # Цена за л
    per_l = re.search(r'за\s+(\d+(?:[.,]\d+)?)\s*л', price_str)
    if per_l:
        qty = float(per_l.group(1).replace(',', '.'))
        return round(price / qty, 2), "л"
    
    # Цена за мл
    per_ml = re.search(r'за\s+(\d+(?:[.,]\d+)?)\s*мл', price_str)
    if per_ml:
        qty = float(per_ml.group(1).replace(',', '.'))
        return round(price / qty, 2), "мл"
    
    # Цена за штуку
    per_pcs = re.search(r'за\s+(\d+(?:[.,]\d+)?)\s*(?:шт|штук|пачку|бутылку|банку|кусочек|упаковку|пакет|пучок|ведро|лоток)', price_str)
    if per_pcs:
        qty = float(per_pcs.group(1).replace(',', '.'))
        return round(price / qty, 2), "шт"
    
    # Если цена указана как "3800 тг/кг"
    if "/кг" in price_str or "тг/кг" in price_str:
        return price, "кг"
    elif "/л" in price_str or "тг/л" in price_str:
        return price, "л"
    elif "/шт" in price_str or "тг/шт" in price_str:
        return price, "шт"
    
    return price, "шт"

def get_quantity_and_unit(qty_str):
    """Возвращает количество и единицу измерения"""
    qty_str = str(qty_str).lower().strip()
    
    # Обработка формата "15 кг 915 г"
    kg_match = re.search(r'(\d+(?:[.,]\d+)?)\s*кг', qty_str)
    g_match = re.search(r'(\d+(?:[.,]\d+)?)\s*г', qty_str)
    
    if kg_match or g_match:
        total = 0
        if kg_match:
            total += float(kg_match.group(1).replace(',', '.'))
        if g_match:
            total += float(g_match.group(1).replace(',', '.')) / 1000
        return round(total, 3), "кг"
    
    # Обработка "2,5 кг"
    simple_kg = re.search(r'(\d+(?:[.,]\d+)?)\s*кг', qty_str)
    if simple_kg:
        return float(simple_kg.group(1).replace(',', '.')), "кг"
    
    # Обработка литров
    l_match = re.search(r'(\d+(?:[.,]\d+)?)\s*л', qty_str)
    if l_match:
        return float(l_match.group(1).replace(',', '.')), "л"
    
    # Обработка штук
    pcs_match = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:шт|штук|пачк|упаковк|бутылк|банк|мешк|кусочк|лоток|ведер|пакет|пучк)', qty_str)
    if pcs_match:
        return float(pcs_match.group(1).replace(',', '.')), "шт"
    
    # Просто число
    num_match = re.search(r'(\d+(?:[.,]\d+)?)', qty_str)
    if num_match:
        return float(num_match.group(1).replace(',', '.')), "шт"
    
    return 1, "шт"

def get_all_purchases():
    """Все 59 позиций с правильным расчётом цены за единицу и количества"""
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
        # Рассчитываем цену за единицу
        price_per_unit, price_unit = parse_price_per_unit(price_str)
        
        # Получаем количество и единицу
        quantity, qty_unit = get_quantity_and_unit(qty_str)
        
        purchases.append({
            "Товар": product,
            "Цена за ед. (₸)": price_per_unit,
            "Количество": quantity,
            "Единица": qty_unit,
            "Сумма (₸)": total,
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
st.caption("Автоматический расчёт цены за кг/л/шт и количества из 59 позиций")

# Загружаем данные
df = load_purchases()

# Кнопка загрузки всех данных
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("📋 ЗАГРУЗИТЬ ВСЕ 59 ПОЗИЦИЙ", use_container_width=True):
        purchases = get_all_purchases()
        for p in purchases:
            p["Дата"] = datetime.now().strftime("%d.%m.%Y")
            p["Поставщик"] = "Замира"
            p["Примечание"] = ""
        
        new_df = pd.DataFrame(purchases)
        st.session_state.loaded_data = new_df
        
        # Сохраняем
        cols_order = ["Дата", "Товар", "Категория", "Цена за ед. (₸)", "Количество", "Единица", "Сумма (₸)", "Поставщик", "Примечание"]
        new_df = new_df[cols_order]
        
        # Объединяем с существующими данными
        if not df.empty:
            df = pd.concat([df, new_df], ignore_index=True)
        else:
            df = new_df
        
        save_purchases(df)
        st.success(f"✅ Загружено {len(purchases)} позиций! Цена за кг/л и количество рассчитаны.")
        st.rerun()

# Боковая панель
with st.sidebar:
    st.markdown("### 📅 Период")
    period = st.radio("Выберите период", ["Сегодня", "Неделя", "Месяц", "Всё время"])
    
    st.markdown("---")
    st.markdown("### ➕ Ручное добавление")
    
    with st.form("add_purchase"):
        product = st.text_input("Товар")
        col1, col2 = st.columns(2)
        with col1:
            price = st.number_input("Цена за ед. (₸)", min_value=0.0, value=0.0, step=10.0)
            qty = st.number_input("Количество", min_value=0.0, value=1.0, step=0.1)
            unit = st.selectbox("Единица", ["кг", "г", "л", "мл", "шт"])
        with col2:
            total = st.number_input("Сумма (₸)", min_value=0.0, value=0.0, step=100.0)
            supplier = st.selectbox("Поставщик", ["Замира", "Магнум", "Метро", "Другое"])
            category = st.selectbox("Категория", list(CATEGORIES.keys()) + ["Прочее"])
        
        if st.form_submit_button("✅ Добавить"):
            if product and (price > 0 or total > 0):
                if total == 0 and price > 0 and qty > 0:
                    total = price * qty
                elif price == 0 and total > 0 and qty > 0:
                    price = total / qty
                
                new_row = pd.DataFrame([{
                    "Дата": datetime.now().strftime("%d.%m.%Y"),
                    "Товар": product,
                    "Категория": category,
                    "Цена за ед. (₸)": round(price, 2),
                    "Количество": qty,
                    "Единица": unit,
                    "Сумма (₸)": round(total, 2),
                    "Поставщик": supplier,
                    "Примечание": ""
                }])
                df = pd.concat([df, new_row], ignore_index=True)
                save_purchases(df)
                st.success("✅ Добавлено!")
                st.rerun()

# Фильтрация
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

# Метрики
c1, c2, c3, c4 = st.columns(4)

if not filtered.empty:
    total_sum = filtered["Сумма (₸)"].sum()
    total_items = len(filtered)
    total_qty = filtered["Количество"].sum() if "Количество" in filtered.columns else 0
    unique_products = filtered["Товар"].nunique()
    
    c1.metric("💰 Общие расходы", f"{total_sum:,.0f} ₸")
    c2.metric("📦 Количество закупок", total_items)
    c3.metric("📊 Общее количество", f"{total_qty:,.1f} ед.")
    c4.metric("🏷️ Уникальных товаров", unique_products)
else:
    c1.metric("💰 Общие расходы", "0 ₸")
    c2.metric("📦 Количество закупок", "0")
    c3.metric("📊 Общее количество", "0")
    c4.metric("🏷️ Уникальных товаров", "0")

st.markdown("---")

# Графики
if not filtered.empty:
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Расходы по дням", "🍽️ По категориям", "📦 По товарам", "📋 Детальный список"])
    
    with tab1:
        if "Дата_парс" in filtered.columns:
            daily = filtered.groupby(filtered["Дата_парс"].dt.date)["Сумма (₸)"].sum().reset_index()
            daily.columns = ["Дата", "Сумма (₸)"]
            if not daily.empty:
                fig = px.line(daily, x="Дата", y="Сумма (₸)", template="plotly_dark", color_discrete_sequence=["#d4a373"])
                fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        by_cat = filtered.groupby("Категория")["Сумма (₸)"].sum().reset_index()
        fig = px.pie(by_cat, values="Сумма (₸)", names="Категория", template="plotly_dark")
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(by_cat.sort_values("Сумма (₸)", ascending=False), use_container_width=True)
    
    with tab3:
        by_product = filtered.groupby("Товар")["Сумма (₸)"].sum().sort_values(ascending=False).head(15).reset_index()
        fig = px.bar(by_product, x="Сумма (₸)", y="Товар", orientation='h', template="plotly_dark", color_discrete_sequence=["#d4a373"])
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=500)
        st.plotly_chart(fig, use_container_width=True)
    
    with tab4:
        display_cols = ["Дата", "Товар", "Категория", "Цена за ед. (₸)", "Количество", "Единица", "Сумма (₸)", "Поставщик"]
        st.dataframe(filtered[display_cols], use_container_width=True, height=400)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            filtered.to_excel(writer, sheet_name="Закупки", index=False)
        st.download_button("📥 Скачать Excel", output.getvalue(), f"закупки_{datetime.now().strftime('%Y-%m-%d')}.xlsx")
else:
    st.info("📭 Нет данных. Нажмите кнопку 'ЗАГРУЗИТЬ ВСЕ 59 ПОЗИЦИЙ'.")

st.caption("💾 Данные сохраняются в 'покупки_кафе.xlsx'")