class GetDataSchema:
    def __init__(
            self,
            pump_wear: float,
            correlation: str,
            coef_correlation: float,
            wht: float,
            sep_method: str,
            sep_method_value: float,
    ):
        self.pump_wear: float = pump_wear
        self.correlation: str = correlation
        self.coef_correlation: float = coef_correlation
        self.wht: float = wht
        self.sep_method: str = sep_method
        self.sep_method_value: float = sep_method_value
