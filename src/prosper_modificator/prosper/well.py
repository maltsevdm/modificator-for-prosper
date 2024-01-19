import enum
from typing import Optional


class Status(enum.Enum):
    close = 0
    open = 1


class WellType(enum.Enum):
    fountain = 'Фонтан'
    esp_with_tms = 'ЭЦН'
    esp_without_tms = 'ЭЦН без ТМС'


class Point:
    def __init__(
            self,
            date: str,
            wellhead_pressure: float,
            annular_pressure: float,
            q_liq: float,
            q_gaz: float,
            wc: float,
            downtime: float,
            wellhead_temperature: float = 10,
            frequency: Optional[float] = None,
            downhole_pressure: Optional[float] = None,
            pump_intake_pressure: Optional[float] = None,
    ):
        self.date: str = date
        self.whp: float = wellhead_pressure
        self.wht: float = wellhead_temperature
        self.ap: float = annular_pressure
        self.q_liq: float = q_liq
        self.q_gaz: float = q_gaz
        self.wc: float = wc
        self.frequency: Optional[float] = frequency
        self.dhp: Optional[float] = downhole_pressure
        self.pip: Optional[float] = pump_intake_pressure
        self.downtime: float = downtime
        self.q_liq_moment = self.calc_q_liq_moment()
        self.gor = self.calc_gor()

    def calc_q_liq_moment(self) -> float:
        return self.q_liq / (1 - self.downtime)

    def calc_gor(self):
        q_liq_moment = self.q_gaz / (self.q_liq_moment * (100 - self.wc / 100))
        return q_liq_moment if q_liq_moment >= 256 else 256


class Well:
    def __init__(
            self,
            mrp: int,
            reservoir_pressure: float,
            prosper_filename: str,
            points: dict,
            avg_gor: float,
            avg_whp: float,
            avg_q_liq: float,
            avg_wc: float,
            avg_frequency: Optional[float] = None,
            avg_pip: Optional[float] = None
    ):
        self.mrp: int = mrp
        self.avg_frequency: Optional[float] = avg_frequency
        self.avg_pip: Optional[float] = avg_pip
        self.avg_gor: Optional[float] = avg_gor
        self.avg_wc: float = avg_wc
        self.avg_q_liq: float = avg_q_liq
        self.avg_whp: float = avg_whp
        self.reservoir_pressure: float = reservoir_pressure
        self.prosper_filename: str = prosper_filename
        self.points: list[Point] = [Point(**point) for point in points]
        self.type = self.define_type()

    def define_type(self):
        for point in self.points:
            if point.frequency:
                break
        else:
            return WellType.fountain

        for point in self.points:
            if point.pip:
                return WellType.esp_with_tms
        else:
            return WellType.esp_without_tms
