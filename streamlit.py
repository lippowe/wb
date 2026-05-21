import streamlit as st
import pandas as pd
import io
import PyPDF2
from datetime import datetime


# 1. Сбор текстов со всех страниц PDF
def get_pdf_pages_content(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    pages_data = []
    for page in pdf_reader.pages:
        text = page.extract_text() or ""
        clean_text = "".join(text.split())
        pages_data.append({"content": clean_text, "page_obj": page})
    return pages_data


# 2. Обработка Excel
def process_xlsx(xlsx_file):
    df_header_raw = pd.read_excel(xlsx_file, header=None, nrows=4)

    try:
        title = str(df_header_raw.iloc[0, 0])
        data_val = str(df_header_raw.iloc[1, 0])
        type_val = str(df_header_raw.iloc[3, 0])
        qty_val = str(df_header_raw.iloc[3, 4])
    except:
        title, data_val, type_val, qty_val = "Лист подбора", "Данные", "Тип", ""

    header_info = {'title': title, 'data': data_val, 'type': type_val, 'qty': qty_val}

    df = pd.read_excel(xlsx_file, skiprows=4)
    cols_to_drop = ['Фото', 'Размер', 'Цвет']
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns], errors='ignore')

    if 'Стикер' in df.columns:
        # Очистка для поиска в PDF (удаляем .0 и пробелы)
        df['Стикер_clean'] = df['Стикер'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\s+', '',
                                                                                                       regex=True)

    # Сортировка по количеству вхождений артикула
    if 'Артикул продавца' in df.columns:
        counts = df['Артикул продавца'].value_counts()
        df['counts'] = df['Артикул продавца'].map(counts)
        df_sorted = df.sort_values(by=['counts', 'Бренд'], ascending=[False, True])

        repeats_mask = df_sorted['Артикул продавца'].duplicated(keep=False)
        df_repeats = df_sorted[repeats_mask].copy()
        df_unique = df_sorted[~repeats_mask].copy()
    else:
        df_repeats, df_unique = pd.DataFrame(), df

    return header_info, df_repeats, df_unique


# 3. Сборка PDF
def create_pdf_output(target_df, pdf_pages):
    if 'Стикер_clean' not in target_df.columns or target_df.empty:
        return None

    writer = PyPDF2.PdfWriter()
    found_count = 0

    for sticker_id in target_df['Стикер_clean']:
        if sticker_id == 'nan': continue
        for page_data in pdf_pages:
            if sticker_id in page_data['content']:
                writer.add_page(page_data['page_obj'])
                found_count += 1
                break

    if found_count == 0:
        return None

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf


# 4. Создание Excel с настройками печати и форматированием
def create_xlsx_output(target_df, header_info):
    if target_df.empty:
        return None

    output = io.BytesIO()
    # Убираем служебные колонки перед сохранением
    cols_to_save = [c for c in target_df.columns if c not in ['counts', 'Стикер_clean']]
    df_final = target_df[cols_to_save]

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, sheet_name='Лист подбора', index=False, startrow=4)

        workbook = writer.book
        worksheet = writer.sheets['Лист подбора']

        # --- ФОРМАТЫ ---
        header_fmt = workbook.add_format({'bold': True, 'font_size': 14})
        bold_fmt = workbook.add_format({'bold': True})
        normal_fmt = workbook.add_format({'font_size': 11})

        # --- ЗАПОЛНЕНИЕ ШАПКИ ---
        worksheet.merge_range('A1:E1', header_info['title'], normal_fmt)
        worksheet.merge_range('A2:I2', header_info['data'], header_fmt)  # Без синего фона
        worksheet.write('A4', header_info['type'], normal_fmt)
        if header_info['qty'] and header_info['qty'] != 'nan':
            worksheet.write('E4', f"Количество: {header_info['qty']}", bold_fmt)

        # --- НАСТРОЙКА КОЛОНОК ---
        sticker_col_idx = -1

        for idx, col_name in enumerate(df_final.columns):
            # Безопасный расчет длины содержимого колонки
            content_len = df_final[col_name].astype(str).str.len().max()
            if pd.isna(content_len): content_len = 0
            max_len = max(content_len, len(col_name)) + 2

            if col_name == 'Наименование':
                worksheet.set_column(idx, idx, min(max_len, 80))  # Максимум 80
                name_col_idx = idx
            elif col_name == 'Стикер':
                worksheet.set_column(idx, idx, max_len)
                sticker_col_idx = idx
            else:
                worksheet.set_column(idx, idx, min(max_len, 40))

        # --- ЖИРНЫЙ ШРИФТ ДЛЯ СТИКЕРА (ПОСЛЕ ПРОБЕЛА) ---
        if sticker_col_idx != -1:
            for row_num, value in enumerate(df_final['Стикер'], start=5):
                val_str = str(value).replace('.0', '').strip()
                if val_str == 'nan': continue

                if " " in val_str:
                    parts = val_str.split(" ", 1)  # Делим только по первому пробелу
                    first_part = parts[0] + " "
                    second_part = parts[1]
                    worksheet.write_rich_string(row_num, sticker_col_idx, first_part, bold_fmt, second_part)
                else:
                    worksheet.write(row_num, sticker_col_idx, val_str)

        # --- НАСТРОЙКИ ПЕЧАТИ ---
        worksheet.set_landscape()  # Альбомная ориентация
        worksheet.set_paper(9)  # А4
        worksheet.set_margins(0.2, 0.2, 0.5, 0.5)  # Узкие поля
        worksheet.fit_to_pages(1, 0)  # Вписать все столбцы на одну страницу по ширине

    output.seek(0)
    return output


def main():
    st.set_page_config(page_title="WB Sticker Pro", layout="wide")
    st.title("📦 Сборка листов подбора WB")

    with st.sidebar:
        pdf_file = st.file_uploader("1. Загрузите PDF со стикерами", type=['pdf'])
        xlsx_file = st.file_uploader("2. Загрузите XLSX файл", type=['xlsx'])

    if pdf_file and xlsx_file:
        with st.spinner('Обработка...'):
            pdf_pages = get_pdf_pages_content(pdf_file)
            header_info, df_repeats, df_unique = process_xlsx(xlsx_file)
            ts = datetime.now().strftime('%H-%M-%S')

            c1, c2 = st.columns(2)
            with c1:
                st.subheader("🔁 Повторы")
                if not df_repeats.empty:
                    pdf_res = create_pdf_output(df_repeats, pdf_pages)
                    xlsx_res = create_xlsx_output(df_repeats, header_info)
                    if pdf_res: st.download_button("📥 Скачать PDF (Повторы)", pdf_res, f"PDF_Rep_{ts}.pdf")
                    if xlsx_res: st.download_button("📥 Скачать Excel (Повторы)", xlsx_res, f"List_Rep_{ts}.xlsx")
                else:
                    st.info("Повторов не найдено")

            with c2:
                st.subheader("🆔 Уникальные")
                if not df_unique.empty:
                    pdf_res = create_pdf_output(df_unique, pdf_pages)
                    xlsx_res = create_xlsx_output(df_unique, header_info)
                    if pdf_res: st.download_button("📥 Скачать PDF (Уникальные)", pdf_res, f"PDF_Uni_{ts}.pdf")
                    if xlsx_res: st.download_button("📥 Скачать Excel (Уникальные)", xlsx_res, f"List_Uni_{ts}.xlsx")
                else:
                    st.info("Уникальных не найдено")


if __name__ == "__main__":
    main()
