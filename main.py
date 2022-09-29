# read excel file to object. send excel object to handler. save data from handler in csv
from parser1 import main as handler1

def main():
    workbook_parsers = {
        r"Задание на парсер\1.xlsx" : handler1,
    }

    for path, func in workbook_parsers.items():
        print(path)
        func(path)




if __name__ == "__main__":
    main()
