import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET
import re
from datetime import datetime, timedelta

# Настройки страницы
st.set_page_config(page_title="Авто-заполнение LOE отчета", layout="wide")
st.title("⚡ Автоматическое заполнение LOE отчета")
st.write("Загрузите файлы, и система сама извлечет продажи, остатки и себестоимость.")

# Функция для корректного чтения специфичных Excel-файлов из 1С
def get_col_index(cell_ref):
    match = re.match(r"([A-Z]+)[0-9]+", cell_ref)
    if not match: return -1
    letters = match.group(1)
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def parse_1c_excel(file_bytes):
    data = []
    try:
        with zipfile.ZipFile(file_bytes, 'r') as z:
            strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    strings = [t.text if t.text is not None else '' for t in root.findall('.//ss:t', ns)]
            
            with z.open('xl/worksheets/sheet1.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'ws': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                rows = root.findall('.//ws:row', ns)
                for row in rows:
                    row_dict = {}
                    for c in row.findall('.//ws:c', ns):
                        r_attr = c.attrib.get('r')
                        col_idx = get_col_index(r_attr) if r_attr else len(row_dict)
                        val_tag = c.find('ws:v', ns)
                        val = val_tag.text if val_tag is not None else ''
                        t = c.attrib.get('t', '')
                        if t == 's' and val.isdigit():
                            idx = int(val)
                            val = strings[idx] if idx < len(strings) else val
                        row_dict[col_idx] = val
                    if row_dict:
                        max_col = max(row_dict.keys())
                        row_data = [row_dict.get(i, '') for i in range(max_col + 1)]
                        data.append(row_data)
    except Exception as e:
        st.error(f"Ошибка чтения файла: {e}")
    return pd.DataFrame(data)

# Интерфейс загрузки файлов
st.subheader("1. Загрузка данных")
col1, col2, col3 = st.columns(3)
with col1:
    base_file = st.file_uploader("1. Общий отчет (loe отчет .xlsx)", type=["xlsx"])
with col2:
    stock_file = st.file_uploader("2. Остатки (Остатки товара .xlsx)", type=["xlsx"])
with col3:
    daily_files = st.file_uploader("3. Ежедневные продажи (17.09.xlsx, 18.09.xlsx...)", type=["xlsx"], accept_multiple_files=True)

if st.button("Сформировать отчет", type="primary", use_container_width=True):
    if not (base_file and stock_file and daily_files):
        st.warning("Пожалуйста, загрузите все необходимые файлы в окна выше!")
    else:
        with st.spinner("Анализирую данные и заполняю отчет..."):
            try:
                # 1. Читаем базовый отчет (шаблон)
                df_main = parse_1c_excel(base_file)
                header = df_main.iloc[0].tolist()
                df_main = df_main[1:].copy()
                df_main.columns = header
                
                new_cols = []
                for c in df_main.columns:
                    try:
                        val = int(c)
                        date_val = datetime(1899, 12, 30) + timedelta(days=val)
                        new_cols.append(date_val.strftime("%d.%m.%Y"))
                    except:
                        new_cols.append(c)
                df_main.columns = new_cols
                
                # 2. Читаем остатки
                df_stock = parse_1c_excel(stock_file)
                stock_data = {}
                header_idx = -1
                for i, row in df_stock.iterrows():
                    if "Артикул" in row.values:
                        header_idx = i
                        break
                        
                if header_idx != -1:
                    headers = df_stock.iloc[header_idx].tolist()
                    for i in range(header_idx + 1, len(df_stock)):
                        row = df_stock.iloc[i].tolist()
                        if len(row) > 1 and row[1]:
                            art = str(row[1]).strip()
                            if not art or art == 'None' or art.startswith('1-100'): continue
                            try:
                                ost_idx = headers.index('Остаток')
                                qty_str = str(row[ost_idx]).strip()
                                if qty_str: stock_data[art] = float(qty_str)
                            except: pass

                # 3. Читаем продажи и себестоимость
                sales_data = {}
                costs_data = {}
                for f in daily_files:
                    date_str = f.name.replace('.xlsx', '.2026') 
                    df_d = parse_1c_excel(f)
                    
                    header_idx = -1
                    for i, row in df_d.iterrows():
                        if "Артикул" in row.values:
                            header_idx = i
                            break
                            
                    if header_idx != -1:
                        headers = df_d.iloc[header_idx].tolist()
                        for i in range(header_idx + 1, len(df_d)):
                            row = df_d.iloc[i].tolist()
                            if len(row) > 1 and row[1]:
                                art = str(row[1]).strip()
                                if not art or art == 'None' or art.startswith(('Итого', 'Средняя', '1.', '2.')): continue
                                
                                try:
                                    kol_idx = headers.index('Кол.')
                                    qty_str = str(row[kol_idx]).strip()
                                    qty = float(qty_str) if qty_str else 0
                                    
                                    seb_idx = headers.index('Себест-сть')
                                    cost_str = str(row[seb_idx]).replace(' ', '').replace(',', '.')
                                    cost = float(cost_str) if cost_str else 0
                                    
                                    if date_str not in sales_data: sales_data[date_str] = {}
                                    sales_data[date_str][art] = qty
                                    costs_data[art] = cost
                                except: pass
                
                # 4. Обновляем главный отчет (ИСПРАВЛЕННЫЙ БЛОК)
                for date_str in sales_data.keys():
                    if date_str not in df_main.columns:
                        df_main[date_str] = None  # Изменили пустую строку на None, чтобы разрешить цифры
                        
                for idx, row in df_main.iterrows():
                    art = str(row['Артикул']).strip()
                    for dt, items in sales_data.items():
                        if art in items:
                            df_main.at[idx, dt] = items[art]
                            
                df_main['Остаток товара'] = df_main['Артикул'].apply(lambda x: stock_data.get(str(x).strip(), None))
                df_main['Себестоимость'] = df_main['Артикул'].apply(lambda x: costs_data.get(str(x).strip(), None))
                
                # 5. Сохраняем и отдаем файл
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_main.to_excel(writer, index=False, sheet_name="Общий отчет")
                output.seek(0)
                
                st.success("✅ Отчет успешно заполнен! Нажмите кнопку ниже, чтобы сохранить его.")
                
                st.download_button(
                    label="📥 Скачать Обновленный_loe_отчет.xlsx",
                    data=output,
                    file_name="Обновленный_loe_отчет.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.subheader("Предварительный просмотр результата:")
                cols_to_show = ['Артикул', 'Наименование'] + list(sales_data.keys()) + ['Остаток товара', 'Себестоимость']
                st.dataframe(df_main[cols_to_show].head(15))
                
            except Exception as e:
                st.error(f"Произошла ошибка при формировании: {e}")
