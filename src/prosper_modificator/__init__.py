import datetime
import os

import json
from typing import Optional

from application.app import App
from excel_file.schemas import InitDataSchema, ResultDataSchema
from excel_file.result_file import ResultFile
from prosper.prosper import Prosper
from prosper.well import Well, WellType, Point


class ProsperModificator:
    def __init__(
            self,
            path_result_file: str,
            prosper_files_directory: str,
            path_init_data: str,
            is_test: bool = False
    ):
        self.is_test = is_test
        self.path_result_file = path_result_file
        self.prosper_files_directory = prosper_files_directory

        self.prosper = Prosper()

        self.data: dict[str, dict[str, Well]] = {}
        with open(path_init_data, encoding='utf-8') as file:
            for field, wells_dict in json.load(file).items():
                self.data[field] = {well: Well(**data) for well, data
                                    in wells_dict.items()}

        fields_data = {field: list(self.data[field].keys())
                       for field in self.data}
        self.app = App(
            on_load=self.on_load,
            on_calc=self.on_calc,
            on_save=self.on_save,
            on_calc_sys=self.on_calc_sys,
            is_test=self.is_test,
            correlations=self.prosper.get_correlations(),
            fields_data=fields_data,
            sep_methods=self.prosper.sep_methods
        )
        self.res_file = ResultFile(path_result_file)

        self.current_well: Optional[Well] = None

        self.app.mainloop()

    def on_load(self):
        field, well = self.app.get_field_well()
        self.current_well = self.data[field][well]

        self.res_file.clear_sheet()
        self.__open_prosper_file()

        correlation = self.prosper.get('ANL.SYS.TubingLabel')
        corr_coef = self.prosper.get(f'ANL.COR.Corr[{{{correlation}}}].A[0]')
        sep_method, sep_method_value, unit = self.prosper.get_sep_method_data()

        data_for_res_file = []
        for point in self.current_well.points:
            data_for_res_file.append(InitDataSchema(**point.__dict__))

        self.app.set_data(
            well_type=self.current_well.type.value,
            mrp=self.current_well.mrp,
            correlation=correlation,
            coef_correlation=corr_coef,
            reservoir_pressure=self.current_well.reservoir_pressure,
            wht=10,
            sep_method=sep_method,
            sep_method_value=sep_method_value,
            unit=unit,
            pump_wear=(self.prosper.get_pump_wear()
                       if self.current_well.type != WellType.fountain
                       else 0),
        )
        if self.current_well.type == WellType.fountain:
            self.app.disable_entry_pump_wear()
        else:
            self.app.enable_entry_pump_wear()
        self.res_file.set_init_data(data_for_res_file)
        print('Данные выгружены.')

    def on_calc(self):
        data_from_app = self.app.get_data()

        new_points = self.res_file.get_init_data()
        self.current_well.points = [Point(**x) for x in new_points]

        self.__open_prosper_file()

        self.prosper.set(f'ANL.COR.Corr[{{{data_from_app.correlation}}}].A[0]',
                         data_from_app.coef_correlation)

        if self.current_well.type != WellType.fountain:
            # Запись метода сепарации и значения
            self.prosper.set_sep_method(data_from_app.sep_method,
                                        data_from_app.sep_method_value)
            # Запись износа насоса
            self.prosper.set('SIN.ESP.Wear', data_from_app.pump_wear)

        if self.current_well.type in [WellType.esp_without_tms,
                                      WellType.fountain]:
            if self.current_well.type == WellType.esp_without_tms:
                # Запись частоты
                self.prosper.set('ANL.WHP.Data[0].Freq',
                                 self.current_well.avg_frequency)
                self.prosper.set('SIN.ESP.Frequency',
                                 self.current_well.avg_frequency)

            dhp, pi, res_points = self.prosper.calculate_using_bhp_from_whp(
                well=self.current_well, wht=data_from_app.wht,
                correlation=data_from_app.correlation
            )
            self.app.set_result(dhp, pi)
        else:
            res_points = self.prosper.calculate_using_tcc(
                well=self.current_well,
                pump_wear=data_from_app.pump_wear,
                correlation=data_from_app.correlation
            )

        points = [ResultDataSchema(
            vlp_ipr_pip=point.vlp_ipr_pip,
            sys_q_liq=point.sys_q_liq,
            sys_pip=point.sys_pip
        ) for point in res_points]
        self.res_file.set_result(points)

    def on_save(self):
        print('Сохраняю...')
        data_from_app = self.app.get_data()

        if self.current_well.type == WellType.esp_with_tms:
            self.__open_prosper_file()

            self.prosper.set(
                f'ANL.COR.Corr[{{{data_from_app.correlation}}}].A[0]',
                data_from_app.coef_correlation
            )
            self.prosper.set('SIN.ESP.Wear', data_from_app.pump_wear)
            self.prosper.set('SIN.ESP.Frequency',
                             self.current_well.avg_frequency)
            self.prosper.set_data_to_vlp_ipr(
                date=datetime.datetime.strftime(datetime.date.today(),
                                                '%d/%m/%Y'),
                whp=self.current_well.avg_whp,
                wc=self.current_well.avg_wc,
                q_liq=self.current_well.avg_q_liq,
                pip=self.current_well.avg_pip,
                gor=self.current_well.avg_gor,
                frequency=self.current_well.avg_frequency,
                wht=data_from_app.wht,
                reservoir_pressure=self.current_well.reservoir_pressure,
                pump_wear=data_from_app.pump_wear
            )

            dhp = self.prosper.calculate_dhp_in_tcc(
                whp=self.current_well.avg_whp,
                wc=self.current_well.avg_wc,
                gor=self.current_well.avg_gor,
                q_liq=self.current_well.avg_q_liq,
                correlation=data_from_app.correlation,
                pip=self.current_well.avg_pip
            )

            if dhp >= self.current_well.reservoir_pressure:
                raise RuntimeError(
                    'Расчётное забойное давление больше пластового!')

            pi = self.prosper.calculate_pi(
                reservoir_pressure=self.current_well.reservoir_pressure,
                wc=self.current_well.avg_wc,
                gor=self.current_well.avg_gor,
                q_liq=self.current_well.avg_q_liq,
                dhp=dhp,
            )

            self.app.set_result(dhp, pi)

        self.prosper.set('ANL.SYS.TubingLabel', data_from_app.correlation)
        self.prosper.set('ANL.VLP.TubingLabel', data_from_app.correlation)

        if not self.is_test:
            self.__save_prosper_file()
        print('Файл Prosper сохранён.')

    def on_calc_sys(self):
        data_from_app = self.app.get_data()
        self.__open_prosper_file()

        if self.current_well != WellType.fountain:
            self.prosper.set_sep_method(data_from_app.sep_method,
                                        data_from_app.sep_method_value)
            self.prosper.set('SIN.ESP.Wear', data_from_app.pump_wear)

        print('Рассчитываю...')
        res_points = []
        for point in self.current_well.points:
            self.prosper.calculate_in_system(
                whp=point.whp, wc=point.wc,
                gor=point.gor, correlation=data_from_app.correlation
            )

            sys_q_liq = self.prosper.get('OUT.SYS.SOL[0].QLIQ')
            res_points.append(ResultDataSchema(sys_q_liq=sys_q_liq))
        self.res_file.set_result(res_points, False)

    def __open_prosper_file(self):
        path = os.path.join(self.prosper_files_directory,
                            self.current_well.prosper_filename)
        self.prosper.open_file(path)

    def __save_prosper_file(self):
        path = os.path.join(self.prosper_files_directory,
                            self.current_well.prosper_filename)
        self.prosper.save_file(path)
