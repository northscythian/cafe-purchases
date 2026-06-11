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

# === ПАРСИНГ ===
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

def parse_docx_purchases(text):
    """Улучшенный парсер для вашего формата DOCX"""
    purchases = []
    lines = text.split('\n')
    
    # Регулярное выражение для поиска строк с данными
    # Ищем: номер, название товара, сумму в конце
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Пропускаем заголовки и разделители
        if line.startswith('№') or line.startswith('--') or 'Товар' in line or 'Цена' in line:
            continue
        
        # Пропускаем строки с только числами или короткие строки
        if len(line) < 10:
            continue
        
        # Ищем сумму в конце строки (последнее число)
        numbers = re.findall(r'(\d{3,6})', line)
        if not numbers:
            continue
        
        total = int(numbers[-1])  # Берём последнее число как сумму
        
        # Извлекаем название товара
        # Удаляем номер в начале
        product_line = re.sub(r'^\d+\s+', '', line)
        # Удаляем сумму в конце
        product_line = re.sub(r'\s+\d{3,6}\s*$', '', product_line)
        # Удаляем цену (числа с тг/₸)
        product_line = re.sub(r'\d+(?:[.,]\d+)?\s*(?:тг|₸)[^\d]*', '', product_line)
        # Удаляем количество (числа с кг/г/л/шт)
        product_line = re.sub(r'\d+(?:[.,]\d+)?\s*(?:кг|г|л|мл|шт|бутылки|пачку|мешка|упаковок|кусочка|лоток|ведер|пакет|пучка|банки)', '', product_line, flags=re.IGNORECASE)
        # Очищаем от лишних пробелов
        product = re.sub(r'\s+', ' ', product_line).strip()
        
        # Если название слишком короткое или пустое — пропускаем
        if not product or len(product) < 3:
            continue
        
        # Определяем количество и единицу
        quantity = 1
        unit = "шт"
        
        # Ищем количество в кг
        kg_match = re.search(r'(\d+(?:[.,]\d+)?)\s*кг', line, re.IGNORECASE)
        g_match = re.search(r'(\d+(?:[.,]\d+)?)\s*г', line, re.IGNORECASE)
        l_match = re.search(r'(\d+(?:[.,]\d+)?)\s*л', line, re.IGNORECASE)
        
        if kg_match:
            quantity = float(kg_match.group(1).replace(',', '.'))
            if g_match:
                quantity += float(g_match.group(1).replace(',', '.')) / 1000
            unit = "кг"
        elif g_match:
            quantity = float(g_match.group(1).replace(',', '.')) / 1000
            unit = "кг"
        elif l_match:
            quantity = float(l_match.group(1).replace(',', '.'))
            unit = "л"
        
        # Определяем цену за единицу
        price_per_unit = 0
        price_match = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:тг|₸)', line)
        if price_match:
            price_val = float(price_match.group(1).replace(',', '.'))
            # Если есть "за ... кг" или "за ... л"
            per_kg = re.search(r'за\s+(\d+(?:[.,]\d+)?)\s*кг', line, re.IGNORECASE)
            per_l = re.search(r'за\s+(\d+(?:[.,]\d+)?)\s*л', line, re.IGNORECASE)
            
            if per_kg:
                qty = float(per_kg.group(1).replace(',', '.'))
                price_per_unit = round(price_val / qty, 2)
            elif per_l:
                qty = float(per_l.group(1).replace(',', '.'))
                price_per_unit = round(price_val / qty, 2)
            else:
                price_per_unit = price_val
        
        purchases.append({
            "Товар": product,
            "Цена за ед.": price_per_unit,
            "Количество": quantity,
            "Единица": unit,
            "Сумма": total
        })
    
    # Удаляем дубликаты по названию товара (оставляем первый)
    unique_purchases = []
    seen_products = set()
    for p in purchases:
        if p["Товар"] not in seen_products:
            seen_products.add(p["Товар"])
            unique_purchases.append(p)
    
    return unique_purchases

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
    df = pd.DataFrame(columns=["Дата", "Товар", "Цена за ед.", "Количество", "Единица", "Сумма", "Поставщик", "Категория", "Примечание"])
    save_data(df)

with st.sidebar:
    st.markdown("### 📅 Период")
    period = st.radio("Показать", ["Сегодня", "Неделя", "Месяц", "Всё время"])
    st.markdown("---")
    
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
    
    st.markdown("### 📎 Загрузить DOCX (Word)")
    uploaded_docx = st.file_uploader("Файл .docx с закупками", type=["docx"])
    
    if uploaded_docx is not None:
        try:
            import docx
            doc = docx.Document(uploaded_docx)
            full_text = []
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    full_text.append(text)
            text = '\n'.join(full_text)
            with st.expander("Просмотр текста"):
                st.text(text[:1000] + "..." if len(text) > 1000 else text)
            if st.button("📊 Распарсить и добавить из DOCX"):
                purchases = parse_docx_purchases(text)
                if purchases:
                    added = 0
                    for p in purchases:
                        new_row = pd.DataFrame([{
                            "Дата": datetime.now().strftime("%d.%m.%Y"),
                            "Товар": p["Товар"],
                            "Цена за ед.": p["Цена за ед."],
                            "Количество": p["Количество"],
                            "Единица": p["Единица"],
                            "Сумма": p["Сумма"],
                            "Поставщик": "Замира",
                            "Категория": detect_category(p["Товар"]),
                            "Примечание": ""
                        }])
                        df = pd.concat([df, new_row], ignore_index=True)
                        added += 1
                    if added > 0:
                        save_data(df)
                        st.success(f"✅ Добавлено {added} позиций из DOCX!")
                        st.rerun()
                    else:
                        st.error("Не удалось распарсить данные")
                else:
                    st.error("Не найдено данных в файле")
        except ImportError:
            st.error("Установите python-docx: pip install python-docx")
        except Exception as e:
            st.error(f"Ошибка: {e}")
    
    st.markdown("---")
    
    if st.button("📋 Загрузить 59 позиций (встроенные)", use_container_width=True):
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
    
    st.markdown("---")
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
        st.subheader("Сводка")
        st.dataframe(r["summary"], use_container_width=True)
        st.subheader("Расходы по категориям")
        st.dataframe(r["by_category"], use_container_width=True)
        st.subheader("Топ-10 товаров")
        st.dataframe(r["top_products"], use_container_width=True)
        st.subheader("Расходы по поставщикам")
        st.dataframe(r["by_supplier"], use_container_width=True)
        st.subheader("ABC-анализ")
        st.dataframe(r["abc"][["Товар", "Сумма", "Доля", "Категория"]], use_container_width=True)
        abc_chart = r["abc"].groupby("Категория")["Сумма"].sum().reset_index()
        if not abc_chart.empty:
            fig = px.pie(abc_chart, values="Сумма", names="Категория", template="plotly_dark", title="ABC-анализ")
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Нажмите 'Сформировать отчёт' для просмотра аналитики")

st.caption("💾 Данные в Google Sheets | Отчёты можно выгрузить в Excel или Google Sheets")
