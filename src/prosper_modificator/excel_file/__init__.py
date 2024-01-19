import xlwings as xw
from xlwings.constants import DeleteShiftDirection

from src.prosper_modificator import InitDataSchema, ResultDataSchema
from src.prosper_modificator.excel_file import formulas
from src.prosper_modificator.excel_file.settings import ResultFileCols


class ResultFile:
    def __init__(self, path):
        self.path = path
        self.wb = xw.Book(path)
        self.ws = self.wb.sheets('Настройка')

    def clear_sheet(self):
        while self.ws.cells(5, 3).value:
            self.ws.range('5:5').api.Delete(DeleteShiftDirection.xlShiftUp)

    def get_init_data(self) -> list[dict]:
        result = []
        for row in range(5, 100):
            result.append(
                {
                    'date': self.ws.cells(row, ResultFileCols.date).value,
                    'q_liq_moment': self.ws.cells(row, ResultFileCols.q_liq_moment).value,
                    'q_liq': self.ws.cells(row, ResultFileCols.q_liq).value,
                    'whp': self.ws.cells(row, ResultFileCols.whp).value,
                    'q_gaz': self.ws.cells(row, ResultFileCols.q_gaz).value,
                    'gor': self.ws.cells(row, ResultFileCols.gor).value,
                    'wc': self.ws.cells(row, ResultFileCols.wc).value,
                    'frequency': self.ws.cells(row, ResultFileCols.frequency).value,
                    'downtime': self.ws.cells(row, ResultFileCols.downtime).value,
                    'pip': self.ws.cells(row, ResultFileCols.pip).value,
                }
            )
        return result


    def set_init_data(
            self, data: list[InitDataSchema],
            disable_screen_updating: bool = True
    ):
        if disable_screen_updating:
            self.wb.app.screen_updating = False

        for i, record in enumerate(data):
            row = i + 5
            self.ws.cells(row, ResultFileCols.date.num).value = record.date
            self.ws.cells(row, ResultFileCols.q_liq.num).value = record.q_liq
            self.ws.cells(row, ResultFileCols.whp.num).value = record.whp
            self.ws.cells(row, ResultFileCols.q_gaz.num).value = record.q_gaz
            self.ws.cells(row, ResultFileCols.wc.num).value = record.wc
            self.ws.cells(
                row, ResultFileCols.downtime.num).value = record.downtime
            self.ws.cells(
                row, ResultFileCols.frequency.num).value = record.frequency
            self.ws.cells(row, ResultFileCols.pip.num).value = record.pip

            self.ws.cells(
                row, ResultFileCols.q_liq_moment.num
            ).value = formulas.get_q_liq_moment_formula(
                let_q_liq=ResultFileCols.q_liq.letter,
                let_downtime=ResultFileCols.downtime.letter,
                row=row
            )

            self.ws.cells(
                row, ResultFileCols.gor.num
            ).value = formulas.get_gor_formula(
                let_q_gaz=ResultFileCols.q_gaz.letter,
                let_q_liq_mom=ResultFileCols.q_liq_moment.letter,
                let_wc=ResultFileCols.wc.letter,
                row=row
            )

            self.do_horiz_align(row)

        if disable_screen_updating:
            self.wb.app.screen_updating = True

    def set_result(
            self, data: list[ResultDataSchema],
            disable_screen_updating: bool = True
    ):
        if disable_screen_updating:
            self.wb.app.screen_updating = False

        for i, record in enumerate(data):
            self.ws.cells(
                i + 5, ResultFileCols.vlp_ipr_pip).value = record.vlp_ipr_pip
            self.ws.cells(
                i + 5, ResultFileCols.sys_q_liq).value = record.sys_q_liq
            self.ws.cells(
                i + 5, ResultFileCols.sys_pip).value = record.sys_pip

            # TODO: Покраска ячеек в зависимости от значения

        if disable_screen_updating:
            self.wb.app.screen_updating = True

    def do_horiz_align(self, row):
        self.ws.range(
            f'{row}:{row}'
        ).api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
