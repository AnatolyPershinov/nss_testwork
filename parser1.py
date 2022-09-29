from openpyxl import load_workbook
from dataclasses import dataclass


@dataclass
class Work:
    index: int
    name: str
    plan_shedule: list
    fact_shedule: list


@dataclass
class Resource:
    index: int
    name: str
    shedule: list


def parser(sheet, days : int):
    # collect information about works and resources
    works: list(Work) = []
    resources: list(Resource) = []

    # part1. parse works
    # work rows contain cell with value "план".   
    for cell in sheet["S"]:
        if cell.value == "план":
            row = sheet[cell.row]
            fact_row = sheet[cell.row+1] # row with fact activity

            # create lists of shedules
            plan_shedule = [c.value for c in row[19:19+days]]
            fact_shedule = [c.value for c in fact_row[19:19+days]]
            works.append(Work(
                index=len(works),
                name=row[1].value+"act",
                plan_shedule=[0 if v is None else v for v in plan_shedule], # replace none to 0
                fact_shedule=[0 if v is None else v for v in fact_shedule]
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
        if cell.value != None and cell.value != "Наименование должностей, профессий":
            resource_rows.append(sheet[cell.row])

    for cell in sheet["B"][end_pos+6:]:
        if cell.value != None and cell.value != "Наименование субподрядной организации" and cell.value != 2:
            resource_rows.append(sheet[cell.row])

    print(resource_rows)

    for resource in resource_rows:
        plan_shedule = [c.value for c in resource[19:19+days]]
        
        resources.append(Resource(
            index=len(resources),
            name=resource[1].value+"_res", 
            shedule=[0 if v is None else v for v in plan_shedule]
        ))

    return works, resources


def save(array, filename):
    pass

def main(path: str):
    workbook = load_workbook(path)
    months = [
        {"name": "Ноябрь", "days" : 30},
    ]

    for month in months:
        ws = workbook.get_sheet_by_name(month["name"])
        parser(ws, days=month["days"])
