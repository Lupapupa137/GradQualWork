import pandas as pd


# правила разделения по постам
def parse_input_table(path):
    df = pd.read_excel(
        path, sheet_name="Лист1", skiprows=19, header=None, usecols="B:I"
    )
    df.columns = [
        "№",
        "Товары (работы, услуги)",
        "Кол-во",
        "Ед.",
        "S1",
        "S",
        "м2",
        "Прим.",
    ]
    df = df[df["№"].apply(lambda x: str(x).isdigit())].reset_index(drop=True)
    return df


def load_post_rules(path):
    posts_df = pd.read_excel(path)
    posts_df.columns = ["Ключ", "Пост"]
    return [
        (str(row["Ключ"]).lower(), str(row["Пост"]))
        for _, row in posts_df.iterrows()
        if pd.notna(row["Ключ"]) and pd.notna(row["Пост"])
    ]
