import os
import streamlit as st
import pandas as pd
import tempfile
from backend.rules import parse_input_table, load_post_rules
from backend.parse_format import split_by_posts_and_export

from backend.database import get_orders, get_order_by_id, insert_order, clean_data


st.set_page_config(page_title="Парсер заказов Stilpark", layout="wide")

st.sidebar.title("Навигация")
page = st.sidebar.radio("Выберите страницу:", ("Парсер заказов", "История заказов"))

if page == "История заказов":
    st.title("📊 История загруженных заказов")

    if st.button("🔄 Обновить историю"):
        try:
            orders = get_orders()

            if not orders:
                st.info("Пока нет загруженных заказов.")
            else:
                orders_data = [o.as_dict() for o in orders]

                df_orders = pd.DataFrame(orders_data)
                st.dataframe(df_orders)

                min_date = df_orders["Дата заказа"].min()
                max_date = df_orders["Дата заказа"].max()
                st.success(f"Доступные данные: с {min_date} по {max_date}")
        except Exception as e:
            st.error(f"Ошибка при получении данных: {e}")
    else:
        st.info("Нажмите кнопку, чтобы загрузить историю заказов.")

    st.stop()


st.title("📦 Stilpark Order Splitter")

st.markdown(
    """
Загрузите входной Excel-файл с заказом, чтобы получить Excel-файлы, разделённые по постам.

**Файл Posts.xlsx** должен лежать рядом с этим приложением и содержать соответствие ключей и номеров постов.
"""
)

uploaded_file = st.file_uploader("Загрузите файл заказа (.xlsx)", type=["xlsx"])

output_folder = "backend/output_tmp"
os.makedirs(output_folder, exist_ok=True)

manager_name = st.text_input("Введите ФИО менеджера (будет в каждом файле):", "")
comments = st.text_input("Введите комментарии (будет в каждом файле):", "")
internal_number = st.text_input("Введите внутренний номер (будет в каждом файле):", "")
delivery_date = st.text_input("Введите дату готовности (будет в каждом файле):", "")

st.markdown(
    "Вы можете загрузить альтернативный файл `Posts.xlsx`. Если не загружен — используется файл по умолчанию."
)
custom_posts_file = st.file_uploader(
    "Загрузите файл соответствий постов (.xlsx)", type=["xlsx"], key="posts"
)

if st.button("🚀 Начать обработку"):
    if not uploaded_file:
        st.warning("Пожалуйста, загрузите файл с заказом.")
    elif not manager_name.strip():
        st.warning("Введите ФИО менеджера.")
    elif not internal_number.strip():
        st.warning("Введите внутренний номер.")
    elif not delivery_date.strip():
        st.warning("Введите дату готовности.")
    else:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name

        # fallback для файла Posts
        if custom_posts_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_posts:
                tmp_posts.write(custom_posts_file.read())
                posts_path = tmp_posts.name
        else:
            posts_path = "backend/Posts.xlsx"

        try:
            df = parse_input_table(tmp_path)
            rules = load_post_rules(posts_path)

            result_files, order_record = split_by_posts_and_export(
                tmp_path,
                df,
                rules,
                output_folder,
                manager_name,
                comments,
                internal_number,
                delivery_date,
            )

            existing_order = get_order_by_id(order_record["order_id"])
            if existing_order:
                st.warning(
                    f"Заказ {order_record['order_id']} уже существует в базе. Запись не выполнена."
                )
            else:
                insert_order(clean_data(order_record))
                st.success(f"Заказ {order_record['order_id']} успешно записан.")

            st.success("Файл успешно сгенерирован!")
            for post_name, file in result_files.items():
                st.download_button(
                    label=f"📥 Скачать {post_name}.xlsx",
                    data=file,
                    file_name=f"{post_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=post_name,
                )
        except Exception as e:
            st.error(f"Ошибка: {e}")
