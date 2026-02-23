# -*- coding: utf-8 -*-
"""
Streamlit web-app — Фестиваль TORONTO
Сторінка 1: Генератор таблиць (Дипломи / Подяки) -> xlsx + PDF
Сторінка 2: Пошук по таблиці -> номер диплому / подяки
"""

import io
import os
import shutil
import sys
import tempfile
import traceback

import pandas as pd
import streamlit as st

# ── Підключаємо логіку з agent.py ────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from agent import build_diplom_table, build_podyaka_table, convert_to_pdf

# ── Налаштування сторінки ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="Фестиваль TORONTO — Таблиці",
    page_icon="🏆",
    layout="centered",
)

# ── Sidebar: навігація ────────────────────────────────────────────────────────
try:
    st.sidebar.image(
        "https://toronto.org.ua/wp-content/uploads/2023/01/toronto-logo.png",
        use_container_width=True,
    )
except Exception:
    pass
st.sidebar.title("Фестиваль TORONTO")
st.sidebar.markdown("toronto.org.ua")
st.sidebar.divider()

page = st.sidebar.radio(
    "Розділ",
    ["📄 Генератор таблиць", "🔍 Пошук по таблиці"],
    index=0,
)


# ══════════════════════════════════════════════════════════════════════════════
# Допоміжна функція генерації (обгортка над agent.py)
# ══════════════════════════════════════════════════════════════════════════════

def _run_generation(
    file_bytes: bytes,
    table_type: str,
    date_str: str,
    name_str: str,
):
    """
    Запускає build_diplom_table або build_podyaka_table у тимчасовій теці.

    ВАЖЛИВО для Windows: tempfile.mkdtemp() може повернути шлях з кирилицею
    (напр. C:\\Users\\Користувач\\AppData\\Local\\Temp\\toronto_xxx).
    openpyxl не може зберегти файл за таким шляхом.
    Рішення: використовуємо папку _tmp поруч з app.py (шлях ASCII).

    Returns:
        (xlsx_bytes, pdf_bytes_or_None, xlsx_filename, error_msg_or_None)
    """
    # Папка для тимчасових файлів — поруч з app.py, шлях без кирилиці
    app_dir = os.path.dirname(os.path.abspath(__file__))
    tmp_dir = os.path.join(app_dir, "_tmp")
    os.makedirs(tmp_dir, exist_ok=True)

    # Унікальний підкаталог для кожного запиту (уникаємо конфліктів)
    import uuid
    session_dir = os.path.join(tmp_dir, uuid.uuid4().hex[:8])
    os.makedirs(session_dir, exist_ok=True)
    tmp_dir = session_dir

    try:
        df = pd.read_excel(io.BytesIO(bytes(file_bytes)), engine="openpyxl")

        # Абсолютний out_name — agent.py збереже файл за цим шляхом
        abs_out_name = os.path.join(tmp_dir, name_str)

        if table_type == "Дипломи":
            raw_path = build_diplom_table(df, date=date_str, out_name=abs_out_name)
        else:
            raw_path = build_podyaka_table(df, date=date_str, out_name=abs_out_name)

        with open(raw_path, "rb") as f:
            xlsx_bytes = f.read()
        xlsx_filename = os.path.basename(raw_path)

        # PDF через LibreOffice (якщо доступний)
        pdf_bytes = None
        try:
            ok = convert_to_pdf(raw_path)
            if ok:
                pdf_path = raw_path.replace(".xlsx", ".pdf")
                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()
        except Exception:
            pass  # PDF недоступний — не критично

        return xlsx_bytes, pdf_bytes, xlsx_filename, None

    except SystemExit:
        # agent.py викликає sys.exit(1) при відсутніх колонках
        try:
            df_check = pd.read_excel(io.BytesIO(bytes(file_bytes)), engine="openpyxl", nrows=0)
            cols = list(df_check.columns)
        except Exception:
            cols = []
        hint = f"\n\nКолонки у вашому файлі: {cols}" if cols else ""
        return None, None, None, (
            "Помилка: обов'язкові колонки не знайдені. "
            "Перевірте що файл містить потрібні колонки (див. довідку вище)."
            + hint
        )
    except Exception as e:
        return None, None, None, f"Помилка: {e}"
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ══════════════════════════════════════════════════════════════════════════════
# СТОРІНКА 1 — ГЕНЕРАТОР ТАБЛИЦЬ
# ══════════════════════════════════════════════════════════════════════════════

def render_generator():
    st.title("Генератор таблиць")
    st.markdown(
        "Завантажте **xlsx-файл учасників**, заповніть параметри "
        "та отримайте готову таблицю у форматі **xlsx** та **PDF**."
    )

    # ── Довідка по колонках ─────────────────────────────────────────────────
    with st.expander("ℹ️ Вимоги до формату xlsx-файлу"):
        st.markdown("""
**Для таблиці Дипломів** файл повинен містити колонки:

| Колонка | Приклад назви |
|---|---|
| ID учасника | `ID` |
| ПІБ учасника | `ПІБ Учасника`, `Artist` |
| Номінація | `Номінація`, `Nomination` |
| Назва роботи | `Назва або опис роботи`, `Title` |
| Laureate | `Laureate` |
| Номер диплому | `№ Диплома`, `Diploma` (опціонально — якщо немає, нумерується 1, 2, 3...) |

**Для таблиці Подяк** файл повинен містити:

| Колонка | Приклад назви |
|---|---|
| ID | `ID` |
| ПІБ керівника | `ПІБ керівника, концертмейстера` |
| Номер подяки | `№Подяки` (опціонально — якщо немає, нумерується автоматично) |
        """)

    st.divider()

    # ── Завантаження файлу ──────────────────────────────────────────────────
    uploaded = st.file_uploader(
        "Файл учасників (.xlsx)",
        type=["xlsx"],
        key="gen_uploader",
        help="Вихідний файл від дизайнера з проставленими номерами дипломів/подяк",
    )

    # Читаємо байти один раз через getvalue() — не потребує seek()
    file_bytes_cached = uploaded.getvalue() if uploaded is not None else None

    if file_bytes_cached is not None:
        # Показуємо попередній перегляд (перші 3 рядки + колонки)
        try:
            df_preview = pd.read_excel(
                io.BytesIO(bytes(file_bytes_cached)), engine="openpyxl", nrows=3
            )
            with st.expander(f"Попередній перегляд файлу ({len(df_preview.columns)} колонок)"):
                st.dataframe(df_preview)
        except Exception as e:
            st.warning(f"Не вдалося прочитати файл: {e}")
            st.code(traceback.format_exc())

    st.divider()

    # ── Параметри ───────────────────────────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        table_type = st.selectbox(
            "Тип таблиці",
            ["Дипломи", "Подяки"],
            help="Дипломи — для учасників, Подяки — для педагогів",
        )
    with col2:
        date_str = st.text_input(
            "Дата відправлення нагород",
            placeholder="напр. 28 лютого або Навесні 2026",
            help="Підставляється у текст опису таблиці",
        )

    name_str = st.text_input(
        "Назва для файлу (без пробілів)",
        placeholder="напр. ТОРОНТО_ЛЮТИЙ_2026",
        help="Використовується у назві вихідного файлу: ДИПЛОМ_ОНЛАЙН_[НАЗВА].xlsx",
    )

    # Замінюємо пробіли на підкреслення щоб не ламати шлях
    name_str_safe = name_str.strip().replace(" ", "_")

    st.divider()

    # ── Кнопка генерації ────────────────────────────────────────────────────
    can_generate = (
        file_bytes_cached is not None
        and date_str.strip() != ""
        and name_str_safe != ""
    )

    if st.button(
        "Згенерувати таблицю",
        type="primary",
        disabled=not can_generate,
        use_container_width=True,
    ):
        with st.spinner("Генерую таблицю... зачекайте"):
            xlsx_b, pdf_b, filename, err = _run_generation(
                file_bytes_cached, table_type, date_str.strip(), name_str_safe
            )

        if err:
            st.error(err)
            # Очищаємо попередній результат
            for k in ("gen_xlsx", "gen_pdf", "gen_name"):
                st.session_state.pop(k, None)
        else:
            st.session_state["gen_xlsx"] = xlsx_b
            st.session_state["gen_pdf"] = pdf_b
            st.session_state["gen_name"] = filename
            st.session_state["gen_type"] = table_type

    # ── Кнопки скачування (зберігаються між перемальовками через session_state)
    if st.session_state.get("gen_xlsx"):
        fname = st.session_state["gen_name"]
        ttype = st.session_state.get("gen_type", "")

        st.success(f"Готово! Таблиця '{fname}' згенерована.")

        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                label="⬇ Завантажити XLSX",
                data=st.session_state["gen_xlsx"],
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl_col2:
            if st.session_state.get("gen_pdf"):
                pdf_name = fname.replace(".xlsx", ".pdf")
                st.download_button(
                    label="⬇ Завантажити PDF",
                    data=st.session_state["gen_pdf"],
                    file_name=pdf_name,
                    mime="application/pdf",
                    use_container_width=True,
                )
            else:
                st.info(
                    "PDF недоступний на цьому сервері.  \n"
                    "Конвертуйте xlsx онлайн: [ilovepdf.com](https://www.ilovepdf.com/excel_to_pdf)"
                )


# ══════════════════════════════════════════════════════════════════════════════
# СТОРІНКА 2 — ПОШУК ПО ТАБЛИЦІ
# ══════════════════════════════════════════════════════════════════════════════

def render_search():
    st.title("Пошук по таблиці")
    st.markdown(
        "Завантажте готову таблицю (xlsx) або вставте дані вручну, "
        "а потім введіть ПІБ учасника або педагога."
    )

    # ── Завантаження даних ──────────────────────────────────────────────────
    tab_upload, tab_paste = st.tabs(
        ["Завантажити xlsx-таблицю", "Вставити дані вручну"]
    )

    with tab_upload:
        sf = st.file_uploader(
            "Готова таблиця (.xlsx)",
            type=["xlsx"],
            key="search_upload",
            help="Таблиця згенерована цим додатком або вручну",
        )
        if sf:
            try:
                # Рядки 1-7 = блок опису (merged), рядок 8 = заголовки
                df_loaded = pd.read_excel(
                    io.BytesIO(bytes(sf.getvalue())),
                    skiprows=7,
                    header=0,
                    engine="openpyxl",
                )
                # Прибираємо повністю порожні рядки (можуть бути після merge)
                df_loaded = df_loaded.dropna(how="all").reset_index(drop=True)
                st.session_state["search_df"] = df_loaded
                st.session_state["search_source"] = sf.name
                st.success(f"Завантажено: {len(df_loaded)} рядків")
            except Exception as e:
                st.error(f"Помилка читання файлу: {e}")

    with tab_paste:
        st.markdown("Формат: `ПІБ,Номер` — кожен учасник з нового рядка")
        pasted = st.text_area(
            "Вставте дані",
            placeholder="Іваненко Марія Петрівна,47\nШевченко Олена,12",
            height=200,
            key="search_paste",
        )
        if st.button("Завантажити вставлені дані"):
            if pasted.strip():
                rows = []
                for line in pasted.strip().splitlines():
                    parts = line.split(",", 1)
                    if len(parts) == 2:
                        rows.append({"Artist": parts[0].strip(), "Номер": parts[1].strip()})
                if rows:
                    st.session_state["search_df"] = pd.DataFrame(rows)
                    st.session_state["search_source"] = "вставлені дані"
                    st.success(f"Завантажено: {len(rows)} рядків")
                else:
                    st.warning("Перевірте формат: кожен рядок має бути 'ПІБ,Номер'")

    # ── Пошук ───────────────────────────────────────────────────────────────
    if "search_df" in st.session_state:
        df = st.session_state["search_df"]
        source = st.session_state.get("search_source", "")

        st.divider()
        st.caption(
            f"Джерело: {source} | "
            f"Рядків: {len(df)} | "
            f"Колонки: {', '.join(str(c) for c in df.columns)}"
        )

        # Автовизначення колонки з іменем
        name_col = None
        for c in df.columns:
            cl = str(c).lower()
            if any(k in cl for k in ["artist", "піб", "ім'я", "имя", "пib"]):
                name_col = c
                break
        if name_col is None and len(df.columns) >= 2:
            name_col = df.columns[1]

        # Автовизначення колонки з номером
        num_col = None
        for c in df.columns:
            cl = str(c).lower()
            if any(k in cl for k in ["диплом", "подяк", "номер", "number", "№"]):
                num_col = c
                break
        if num_col is None:
            num_col = df.columns[-1]

        query = st.text_input(
            "Введіть ім'я учасника або педагога",
            placeholder="напр. Іваненко або Шевченко",
            key="search_query",
        )

        if query.strip():
            mask = (
                df[name_col]
                .astype(str)
                .str.contains(query.strip(), case=False, na=False)
            )
            results = df[mask].reset_index(drop=True)

            if len(results) == 0:
                st.warning(f"Нічого не знайдено за запитом: «{query}»")

            elif len(results) == 1:
                row = results.iloc[0]
                name_val = str(row[name_col])
                num_val  = str(row[num_col])

                st.success("Знайдено!")
                st.metric(
                    label=name_val,
                    value=f"№ {num_val}",
                )

                # Повний рядок для довідки
                with st.expander("Детальна інформація"):
                    for col in results.columns:
                        val = str(row[col])
                        if val and val != "nan":
                            st.write(f"**{col}:** {val}")

            else:
                st.info(f"Знайдено {len(results)} збігів:")
                # Показуємо тільки ключові колонки
                display_cols = [name_col, num_col]
                for c in df.columns:
                    if c not in display_cols and str(c).lower() not in ["id", "unnamed: 0"]:
                        display_cols.append(c)
                        if len(display_cols) >= 5:
                            break
                st.dataframe(
                    results[display_cols],
                    hide_index=True,
                )

        else:
            # Показуємо всю таблицю якщо запит порожній
            with st.expander("Переглянути всі записи"):
                st.dataframe(df, hide_index=True)

    st.sidebar.divider()
    st.sidebar.markdown(
        "**Підказка:** Спочатку згенеруйте таблицю на сторінці "
        "«Генератор», потім завантажте її тут для пошуку."
    )


# ══════════════════════════════════════════════════════════════════════════════
# РОУТЕР
# ══════════════════════════════════════════════════════════════════════════════

if page == "📄 Генератор таблиць":
    render_generator()
else:
    render_search()
