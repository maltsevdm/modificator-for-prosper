from typing import Union, Optional

from openserver import OpenServer

from src.prosper_modificator import Well
from src.prosper_modificator.prosper.correlation import Correlations


class ResultPoint:
    def __init__(
            self,
            sys_q_liq: Optional[float] = None,
            sys_pip: Optional[float] = None,
            vlp_ipr_pip: Optional[float] = None
    ):
        self.sys_q_liq: Optional[float] = sys_q_liq
        self.sys_pip: Optional[float] = sys_pip
        self.vlp_ipr_pip: Optional[float] = vlp_ipr_pip


class ProsperInterface:
    def __init__(self):
        self.c: OpenServer = OpenServer()
        self.c.connect()
        self.correlations = Correlations
        self.sep_methods = ['Эффективность сепарации', 'Диаметр штуцера']

    def get(self, query: str):
        return self.c.DoGet('PROSPER.' + query)

    def cmd(self, query: str):
        self.c.DoCmd('PROSPER.' + query)

    def set(self, query: str, value: Union[str, float, int]):
        self.c.DoSet('PROSPER.' + query, value)

    def open_file(self, path):
        print('Opening file...')
        self.cmd(f'OPENFILE({path})')
        print('File opened.')

    def save_file(self, path):
        self.cmd(f'PROSPER.SAVEFILE({path})')
        print('File saved.')

    def refresh(self):
        self.cmd('REFRESH')

    def set_anl_vmt_data(self, index: int, param: str, value):
        self.set(f'ANL.VMT.Data[{index}].{param}', value)

    def set_anl_whp_data(self, index: int, param: str, value):
        self.set_anl_whp(f'Data[{index}].{param}', value)

    def set_sin_esp(self, param: str, value):
        self.set(f'SIN.ESP.{param}', value)

    def set_anl_sys(self, param: str, value):
        self.set(f'ANL.SYS.{param}', value)

    def set_anl_tcc(self, param: str, value):
        self.set(f'ANL.TCC.{param}', value)

    def set_anl_whp(self, param: str, value):
        self.set(f'ANL.WHP.{param}', value)

    def set_sin_ipr_single(self, param: str, value):
        self.set(f'SIN.IPR.Single.{param}', value)

    def calculate_ipr(self):
        self.cmd('IPR.Calc')

    def get_pump_depth(self):
        return self.get('SIN.ESP.Depth')

    def get_sep_method(self) -> str:
        sep_method_index = self.get('SIN.ESP.SepMethod')
        return self.sep_methods[sep_method_index]

    def get_sep_method_data(self) -> tuple[str, float, str]:
        sep_method_index = self.get('SIN.ESP.SepMethod')
        sep_method = self.sep_methods[sep_method_index]
        if sep_method_index == 0:
            unit = self.get('SIN.ESP.Efficiency.Unitname')
            sep_method_value = self.get('SIN.ESP.Efficiency')
        else:
            unit = self.get('SIN.ESP.SepPortDiam.Unitname')
            sep_method_value = self.get('SIN.ESP.SepPortDiam')
        return sep_method, sep_method_value, unit

    def set_sep_method(self, name: str, value: float):
        index = self.sep_methods.index(name)
        self.set_sin_esp('SepMethod', index)
        if index == 0:
            self.set_sin_esp('Efficiency', value)
        elif index == 1:
            self.set_sin_esp('RgsDPortMethod', 1)
            self.set_sin_esp('SepPortDiam', value)

    @staticmethod
    def get_correlations() -> list[str]:
        return Correlations.get_correlations()


class Prosper(ProsperInterface):
    def calculate_pi(
            self, reservoir_pressure: float, wc: float, gor: float,
            q_liq: float, dhp: float
    ):
        # Расчёта PI в секции IPR
        self.set_sin_ipr_single('IprMethod', 1)
        self.set_sin_ipr_single('Pres', reservoir_pressure)
        self.set_sin_ipr_single('Tres', 104)
        self.set_sin_ipr_single('WC', wc)
        self.set_sin_ipr_single('totgor', gor)
        self.set_sin_ipr_single('Qtest', q_liq)
        self.set_sin_ipr_single('Ptest', dhp)
        self.refresh()
        self.calculate_ipr()

        pi = self.get('Sin.IPR.Single.PINSAV')

        self.set_sin_ipr_single('IprMethod', 0)
        self.set_sin_ipr_single('Pindex', pi)
        self.refresh()
        self.calculate_ipr()
        return pi

    def set_data_to_vlp_ipr(
            self,
            date: str, whp: float, wc: float, q_liq: float, pip: float,
            gor: float, frequency: float, wht: float,
            reservoir_pressure: float, pump_wear: float
    ):
        row_num = self.get('ANL.VMT.Data.Count')

        # Запись данных в секцию VLP/IPR
        self.set_anl_vmt_data(row_num, 'Date', date.replace('.', '/'))
        self.set_anl_vmt_data(row_num, 'THpres', whp)
        self.set_anl_vmt_data(row_num, 'THtemp', wht)
        self.set_anl_vmt_data(row_num, 'WC', wc)
        self.set_anl_vmt_data(row_num, 'Rate', q_liq)

        pump_depth = self.get_pump_depth()

        self.set_anl_vmt_data(row_num, 'Gdepth', pump_depth)
        self.set_anl_vmt_data(row_num, 'Gpres', pip)
        self.set_anl_vmt_data(row_num, 'Pres', reservoir_pressure)
        self.set_anl_vmt_data(row_num, 'GOR', gor)
        self.set_anl_vmt_data(row_num, 'GORfree', 0)
        self.set_anl_vmt_data(row_num, 'Freq', frequency)
        self.set_anl_vmt_data(row_num, 'Wear', pump_wear)
        self.set_anl_vmt_data(row_num, 'PIP', pip)
        self.set_anl_vmt_data(row_num, 'PDP', 2)

    def set_correlation_in_tcc(self, correlation_index: int):
        self.set_anl_tcc('CORR [$]', 0)
        self.set_anl_tcc(f'CORR[{correlation_index}]', 1)

    def set_data_in_tcc(
            self, whp: float, wc: float, gor: float, q_liq: float,
            correlation: str, pip_fact: Optional[float] = None
    ):
        # Запись данных в секцию TCC
        self.set_anl_tcc('Pres', whp)
        self.set_anl_tcc('WC', wc)
        self.set_anl_tcc('GOR', gor)
        self.set_anl_tcc('GORFree', 0)
        self.set_anl_tcc('Rate', q_liq)

        for i in range(0, 10):
            if self.get(f'ANL.TCC.Comp[{i}].Msd') is None:
                break
            else:
                self.set_anl_tcc(f'Comp[{i}].Msd', '')
                self.set_anl_tcc(f'Comp[{i}].Prs', '')

        pump_depth = self.get_pump_depth()

        self.set_anl_tcc('Comp[0].Msd', pump_depth)
        self.set_anl_tcc('Comp[0].Prs', pip_fact)

        # Проставление флажка около нужной корреляции в TCC
        corr_index = self.correlations.get_index_by_name(correlation)
        self.set_correlation_in_tcc(corr_index)
        self.refresh()
        self.cmd('ANL.TCC.CALC')

    def calculate_dhp_in_bhp_from_whp(
            self, date: str, q_liq: float, whp: float, wht: float,
            gor: float, wc: float, correlation: str
    ):
        # Запись в BHP from WHP
        self.set_anl_whp_data(0, 'Time', date)
        self.set_anl_whp_data(0, 'Rate', q_liq)
        self.set_anl_whp_data(0, 'WHP', whp)
        self.set_anl_whp_data(0, 'WHT', wht)
        self.set_anl_whp_data(0, 'GASF', gor)
        self.set_anl_whp_data(0, 'WATF', wc)
        self.set_anl_whp('TubingLabel', correlation)
        self.refresh()
        self.cmd('ANL.WHP.CALC')
        return self.get('ANL.WHP.Data[0].BHP')

    def calculate_dhp_in_tcc(
            self, whp: float, wc: float, gor: float, q_liq: float,
            correlation: str, pip: float
    ):
        self.set_data_in_tcc(
            whp=whp, wc=wc, gor=gor, q_liq=q_liq,
            correlation=correlation, pip_fact=pip
        )

        corr_index = self.correlations.get_index_by_name(correlation)
        value_regime = list(self.get(f'OUT.TCC.RES[{corr_index}][0][$].Regime'))

        index_bhp = value_regime.index('Bottomhole')
        dhp = self.get(f'OUT.TCC.RES[{corr_index}][0][{index_bhp}].Pres')

        return dhp

    def calculate_pip_in_tcc(
            self, whp: float, wc: float, gor: float, q_liq: float,
            correlation: str, pip: float
    ):
        self.set_data_in_tcc(
            whp=whp, wc=wc, gor=gor, q_liq=q_liq,
            correlation=correlation, pip_fact=pip
        )

        corr_index = self.correlations.get_index_by_name(correlation)
        value_regime = list(self.get(f'OUT.TCC.RES[{corr_index}][0][$].Regime'))

        try:
            index_pip = value_regime.index('Pump Intake')
        except ValueError:
            return -1
        return self.get(f'OUT.TCC.RES[{corr_index}][0][{index_pip}].Pres')

    def get_pump_wear(self) -> float:
        return self.get('SIN.ESP.Wear')



    def calculate_using_tcc(
            self, well: Well, pump_wear: float, correlation: str,
    ):
        res_points = []
        for point in well.points:
            self.set_sin_esp('Frequency', point.frequency)

            # Запись данных в VLP/IPR
            self.set_data_to_vlp_ipr(
                date=point.date,
                whp=point.whp,
                wc=point.wc,
                q_liq=point.q_liq_moment,
                pip=point.pip,
                gor=point.gor,
                frequency=point.frequency,
                wht=point.wht,
                reservoir_pressure=well.reservoir_pressure,
                pump_wear=pump_wear
            )

            # Получение давления на приёме
            pip = self.calculate_pip_in_tcc(
                whp=point.whp,
                wc=point.wc,
                gor=point.gor,
                q_liq=point.q_liq,
                correlation=correlation,
                pip=point.pip
            )
            res_points.append(ResultPoint(vlp_ipr_pip=pip))

        return res_points

    def calculate_using_bhp_from_whp(
            self, well: Well, wht: float, correlation: str,
    ) -> tuple[float, float, list[ResultPoint]]:
        dhp = self.calculate_dhp_in_bhp_from_whp(
            date='', q_liq=well.avg_q_liq, whp=well.avg_whp, wht=wht,
            gor=well.avg_gor, wc=well.avg_wc, correlation=correlation
        )

        if dhp >= well.reservoir_pressure:
            raise RuntimeError(
                'Расчётное забойное давление больше пластового!')

        pi = self.calculate_pi(well.reservoir_pressure, well.avg_wc,
                               well.avg_gor, well.avg_q_liq, dhp)

        res_points = []
        for point in well.points:
            self.calculate_in_system(
                whp=point.whp, wc=point.wc,
                gor=point.gor, correlation=correlation
            )

            res_points.append(ResultPoint(
                sys_q_liq=self.get('OUT.SYS.SOL[0].QLIQ'),
                sys_pip=self.get('OUT.SYS.SOL[0].PIP')
            ))

        return dhp, pi, res_points

    def calculate_in_system(
            self, whp: float, wc: float, gor: float, correlation: str
    ):
        # Расчёт в System
        self.set_anl_sys('Pres', whp)
        self.set_anl_sys('WC', wc)
        self.set_anl_sys('GOR', gor)
        self.set_anl_sys('TubingLabel', correlation)
        self.cmd('ANL.SYS.Calc')
