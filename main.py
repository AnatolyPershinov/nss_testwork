# manage parsers
from parsers.parser1 import main as handler1


def main():
    workbook_parsers = {
        r"Задание на парсер\1.xlsx" : handler1,
    }

    for path, func in workbook_parsers.items():
        print(path)
        func(path)


if __name__ == "__main__":
    main()
