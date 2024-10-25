import re


if __name__ == "__main__":
    date_pattern = r'\d{4}/\d{1,2}/\d{1,2}'
    date_text = "2021/11/15"
    print(re.match(date_pattern, date_text) is None)
    date_array = date_text.split("/")
    print(date_array)
