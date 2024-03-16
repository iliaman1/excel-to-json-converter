import datetime
import json
from dataclasses import dataclass
from pprint import pprint

import openpyxl


@dataclass
class Month:
    d_201: float
    tax: float
    b_610: float
    b_600: float
    B_620: float
    c_650: float
    o_660: float
    mat_pom: float


@dataclass
class RawData:
    number: int
    full_name: str
    passport_number: str
    personal_number: str
    address: str
    months: list[Month]

    def generate_tar4(self) -> list[dict]:
        tar4 = []
        for index in range(12):
            tar4.append({
                'nmonth': index + 1,
                'nsummonth': self.months[index].d_201,
                'tar4sum': [{'ncode': 201, 'nsum': self.months[index].d_201}]
            })
        return tar4

    def generate_tar7(self) -> list[dict]:
        tar7 = []
        for index in range(12):
            tar7.append({
                'nmonth': index + 1,
                'nsummonth': self.months[index].b_600,
                'tar7sum': [{'ncode': 600, 'nsum': self.months[index].b_600}]
            })
        return tar7

    def generate_tar14(self) -> list[dict]:
        tar14 = []
        for index in range(12):
            tar14.append({
                'nmonth': index + 1,
                'nsumdiv': 0,
                'nsumt': self.months[index].tax
            })
        return tar14

    def serialize_to_dict(self):
        fio = self.full_name.split()
        if len(fio) < 3:
            fio += ['', '', '']
        return {
            'docagentinfo': {
                'cln': self.personal_number,
                'cstranf': '112',  # 2.5 Код страны гражданства (подданства)
                'cvdoc': '01',  # 2.6 Код документа, удостоверяющего личность
                'nrate': 13,  # 3. Размер ставки подоходного налога с физических лиц, проценты
                'vfam': fio[0],
                'vname': fio[1],
                'votch': fio[2]
            },
            'nsumstand': sum([nalog.b_600 for nalog in self.months]),
            'ntsumbank': 0,
            'ntsumcalcincome': sum([nalog.tax for nalog in self.months]),
            'ntsumcalcincomediv': 0,
            'ntsumexemp': 0,
            'ntsumincome': sum([nalog.d_201 for nalog in self.months]),
            'ntsumnotcalc': 0,
            'ntsumprof': 0,
            'ntsumprop': 0,
            'ntsumsec': 0,
            'ntsumsoc': 0,
            'ntsumtrust': 0,
            'ntsumwithincome': 0,
            'ntsumwithincomediv': 0,
            'tar14': self.generate_tar14(),
            'tar4': self.generate_tar4(),
            'tar7': self.generate_tar7()
        }


def create_raw_data_list(excel_filename: str) -> list[RawData]:
    data: list[RawData] = []
    wb = openpyxl.load_workbook(excel_filename)
    sheet = wb.active
    max_rows = sheet.max_row

    for xlsx_row in range(5, max_rows + 1, 9):
        months = []
        for month in range(2, 14):
            months.append(
                Month(
                    sheet.cell(row=xlsx_row + 1, column=month).value if sheet.cell(row=xlsx_row + 1,
                                                                                   column=month).value else 0,
                    sheet.cell(row=xlsx_row + 2, column=month).value if sheet.cell(row=xlsx_row + 2,
                                                                                   column=month).value else 0,
                    sheet.cell(row=xlsx_row + 3, column=month).value if sheet.cell(row=xlsx_row + 3,
                                                                                   column=month).value else 0,
                    sheet.cell(row=xlsx_row + 4, column=month).value if sheet.cell(row=xlsx_row + 4,
                                                                                   column=month).value else 0,
                    sheet.cell(row=xlsx_row + 5, column=month).value if sheet.cell(row=xlsx_row + 5,
                                                                                   column=month).value else 0,
                    sheet.cell(row=xlsx_row + 6, column=month).value if sheet.cell(row=xlsx_row + 6,
                                                                                   column=month).value else 0,
                    sheet.cell(row=xlsx_row + 7, column=month).value if sheet.cell(row=xlsx_row + 7,
                                                                                   column=month).value else 0,
                    sheet.cell(row=xlsx_row + 8, column=month).value if sheet.cell(row=xlsx_row + 8,
                                                                                   column=month).value else 0
                )
            )

        data.append(
            RawData(
                sheet.cell(row=xlsx_row, column=1).value,
                sheet.cell(row=xlsx_row, column=2).value,
                sheet.cell(row=xlsx_row, column=6).value,
                sheet.cell(row=xlsx_row, column=8).value,
                sheet.cell(row=xlsx_row, column=10).value,
                months
            )
        )
    return data


rawdata = create_raw_data_list('доход2023.xlsx')


def serialize_to_json(rawdata: list[RawData]):
    current_time = datetime.datetime.now()
    formatted_time = current_time.strftime("%Y-%m-%dT%H:%M:%S")
    return {'pckagent': {
        'docagent': [person.serialize_to_dict() for person in rawdata],
        'pckagentinfo': {
            'dcreate': formatted_time,
            'ngod': current_time.year,
            'nmns': 741,
            'nmnsf': 741,
            'ntype': 1,
            'vexec': 'Буйко Т.С.',
            'vphn': '72-30-50',
            'vunp': '700069297'
        }}}


def generate_filename(unp, form_type, department_code, part_number=None):
    current_time = datetime.datetime.now()
    formatted_time = current_time.strftime("%Y%m%d%H%M%S")
    filename = f"D{unp}_{current_time.year}_{form_type}_{department_code}_{formatted_time}"
    if part_number is not None:
        filename += f"_{part_number:04d}"
    filename += ".json"
    return filename


def batch(iterable, n=1):
    l = len(iterable)
    for ndx in range(0, l, n):
        yield iterable[ndx:min(ndx + n, l)]


def make_files(data: list[RawData]):
    unp = 700069297

    part_counts = len(data) // 200
    if part_counts == 0:
        data = serialize_to_json(data)
        with open(
                f'gen_json/{generate_filename(unp, 1, 0)}',
                'w',
                encoding='utf-8'
        ) as write_file:
            json.dump(data, write_file, indent=4, ensure_ascii=False)
        return

    for part, group in enumerate(batch(data, 200)):
        partdata = serialize_to_json(group)
        with open(
                f'gen_json/{generate_filename(unp, 1, 0, part + 1)}',
                'w',
                encoding='utf-8'
        ) as write_file:
            json.dump(partdata, write_file, indent=4, ensure_ascii=False)


make_files(rawdata)
