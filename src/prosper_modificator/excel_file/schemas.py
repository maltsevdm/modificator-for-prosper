from typing import Optional


class InitDataSchema:
    def __init__(
            self,
            date: str,
            q_liq: float,
            whp: float,
            q_gaz: float,
            wc: float,
            downtime: float,
            frequency: Optional[float] = None,
            pip: Optional[float] = None,
            **kwargs
    ):
        self.date: str = date
        self.q_liq: float = q_liq
        self.whp: float = whp
        self.q_gaz: float = q_gaz
        self.wc: float = wc
        self.downtime: float = downtime
        self.frequency: Optional[float] = frequency
        self.pip: Optional[float] = pip


class ResultDataSchema:
    def __init__(
            self,
            vlp_ipr_pip: Optional[float] = None,
            sys_q_liq: Optional[float] = None,
            sys_pip: Optional[float] = None,
            **kwargs
    ):
        self.vlp_ipr_pip: Optional[float] = vlp_ipr_pip
        self.sys_q_liq: Optional[float] = sys_q_liq
        self.sys_pip: Optional[float] = sys_pip
