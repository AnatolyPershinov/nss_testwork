import csv

from typing import List
from dataclasses import dataclass
from openpyxl import load_workbook


@dataclass
class Work:
    index: int
    name: str
    plan_schedule: list
    fact_schedule: list

    def get_dict(self):
        result = {"Название работы": self.name}
        header = [x for x in range(1, len(self.plan_schedule)+1)]
        result.update(dict(zip([str(day) +" plan" for day in header], self.plan_schedule)))
        result.update(dict(zip([str(day) +" fact" for day in header], self.fact_schedule)))
        return result

@dataclass
class Resource:
    index: int
    name: str
    schedule: list
    
    def get_dict(self):
        result = {"Название ресурса": self.name}
        header = [x for x in range(1, len(self.schedule)+1)]
        result.update(dict(zip([str(day) for day in header], self.schedule)))
        return result


def parser(sheet, days : int) ->  tuple[list[Work], list[Resource]]:
    # collect information about works and resources
    works: List[Work] = []
    resources: List[Resource] = []

    # part1. parse works
    # work rows contain cell with value "план"   
    for cell in sheet["S"]:
        if cell.value != "план":
            continue

        row = sheet[cell.row]
        fact_row = sheet[cell.row+1] # row with fact activity

        # create lists of schedules
        plan_schedule = [c.value for c in row[19:19+days]]
        fact_schedule = [c.value for c in fact_row[19:19+days]]
        works.append(Work(
            index=len(works),
            name=row[1].value+"_act",
            plan_schedule=[0 if v is None else v for v in plan_schedule], # replace none to 0
            fact_schedule=[0 if v is None else v for v in fact_schedule]
        ))
        
    # part2. parse resources
    start_pos = 0
    end_pos = 0
    resource_rows = []

    for cell in sheet["B"]:
        if cell.value == "Наименование и марка техники (механизма), оборудования":
            start_pos = cell.row + 2
        elif cell.value == "Наименование субподрядной организации":
            end_pos = cell.row-4
            break
    
    for cell in sheet["B"][start_pos:end_pos]:
        if cell.value is not None and cell.value != "Наименование должностей, профессий":
            resource_rows.append(sheet[cell.row])

    for cell in sheet["B"][end_pos+6:]:
        if cell.value is not None and cell.value != "Наименование субподрядной организации" and cell.value != 2:
            resource_rows.append(sheet[cell.row])

    for resource in resource_rows:
        plan_schedule = [c.value for c in resource[19:19+days]]
        
        resources.append(Resource(
            index=len(resources),
            name=resource[1].value+"_res", 
            schedule=[0 if v is None else v for v in plan_schedule]
        ))

    return works, resources


def save_to_csv(filename, array):
    with open(filename, "w", encoding="UTF-8", newline='\r\n') as csvfile:
        header = array[0].get_dict().keys()
        writer = csv.DictWriter(f=csvfile, fieldnames=header, lineterminator="\n")
        writer.writeheader()
        for k in array:
            writer.writerow(k.get_dict())


def main(path: str):
    workbook = load_workbook(path)
    months = [
        {"name": "Ноябрь", "days" : 30},
    ]

    for month in months:
        ws = workbook.get_sheet_by_name(month["name"])
        works, resources = parser(ws, days=month["days"])
        save_to_csv("activites.csv", works)
        save_to_csv("resources.csv", resources)

