from typing import Optional

import tkinter as tk
from tkinter import ttk

from src.prosper_modificator.app.fonts import FONT_LABELS, FONT_TEXT
from src.prosper_modificator.app.schemas import GetDataSchema


class AskRowForm(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title('Ввод строки')
        self.attributes("-topmost", True)
        self.row_save = tk.StringVar()

        # Label Строка
        lbl_row = tk.Label(self, text='Строка:', font=FONT_LABELS)
        lbl_row.grid(row=1, column=0, padx=1, pady=1, sticky='e')

        # Entry строка сохранения
        ent_row = tk.Entry(self, width=23, relief=tk.GROOVE,
                           textvariable=self.row_save, borderwidth=2)
        ent_row.grid(row=1, column=1, padx=1, pady=1)
        ent_row.focus()

        # Кнопка "Сохранить"
        btn_save = ttk.Button(self, text='Сохранить', command=self.destroy)
        btn_save.grid(row=2, column=0, padx=1, pady=1)

    def open(self):
        self.grab_set()
        self.wait_window()
        return self.row_save.get()


class App(tk.Tk):
    def __init__(
            self, correlations: list[str], fields_data: dict[str, list],
            on_load, on_calc, on_save, on_calc_sys, sep_methods: list,
            is_test: bool = False
    ):
        super().__init__()

        self.correlations: list[str] = correlations
        self.fields_data: dict[str, list] = fields_data
        self.is_test: bool = is_test
        self.sep_methods = sep_methods

        self.title('Настройка Prosper')
        self.attributes("-topmost", True)

        # Label Месторождение
        lbl_field = tk.Label(self, text='Месторождение:', font=FONT_LABELS)
        lbl_field.grid(row=1, column=0, padx=1, pady=1, sticky='e')

        # ComboxBox месторождений
        fields = list(self.fields_data.keys())
        self.combo_field = ttk.Combobox(self, values=fields)
        self.combo_field.grid(row=1, column=1, padx=1, pady=1)
        self.combo_field.bind("<<ComboboxSelected>>", self.__update_combo_well)

        # Label Скважина
        lbl_well = tk.Label(self, text='Скважина:', font=FONT_LABELS)
        lbl_well.grid(row=2, column=0, padx=1, pady=1, sticky='e')

        # ComboxBox скважин
        self.combo_well = ttk.Combobox(self, values=[])
        self.combo_well.grid(row=2, column=1, padx=1, pady=1)

        # Кнопка "Выгрузить"
        btn_load = ttk.Button(self, text='Выгрузить', command=on_load)
        btn_load.grid(row=3, column=1, padx=1, pady=1)

        row = 4
        for text in ['Тип:', 'МРП, дней:', 'Пласт. давление, атм:',
                     'Корреляция:',
                     'Коэф. корреляции:', 'Износ насоса:',
                     'Темп. на буфере, °С:',
                     'Метод сепарации газа:']:
            lbl = tk.Label(self, text=text, font=FONT_LABELS)
            lbl.grid(row=row, column=0, padx=1, pady=1, sticky='e')
            row += 1

        self.lbl_gsm = tk.Label(self, text='Диаметр, мм', font=FONT_LABELS)
        self.lbl_gsm.grid(row=12, column=0, padx=1, pady=1, sticky='e')

        # Тип текст
        self.lbl_type_text = tk.Label(self, text='', font=FONT_TEXT)
        self.lbl_type_text.grid(row=4, column=1, padx=1, pady=1)

        # МРП текст
        self.lbl_mrp_text = tk.Label(self, text='', font=FONT_TEXT)
        self.lbl_mrp_text.grid(row=5, column=1, padx=1, pady=1)

        # Пластовое давление текст
        self.lbl_reser_pres_text = tk.Label(self, text='', font=FONT_TEXT)
        self.lbl_reser_pres_text.grid(row=6, column=1, padx=1, pady=1)

        # Combobox корреляций
        self.combo_corr = ttk.Combobox(self, values=correlations)
        self.combo_corr.grid(row=7, column=1, padx=1, pady=1)

        # Entry коэф. корреляции
        self.ent_coef_corr = tk.Entry(self, width=23, relief=tk.GROOVE,
                                      borderwidth=2)
        self.ent_coef_corr.grid(row=8, column=1, padx=1, pady=1)

        # Entry износ насоса
        self.ent_pump_wear = tk.Entry(self, width=23, relief=tk.GROOVE,
                                      borderwidth=2)
        self.ent_pump_wear.grid(row=9, column=1, padx=1, pady=1)

        # Entry темп. буфера
        self.ent_whp = tk.Entry(self, width=23, relief=tk.GROOVE,
                                borderwidth=2)
        self.ent_whp.grid(row=10, column=1, padx=1, pady=1)

        # Combobox методов сепарации газа
        self.combo_gsm = ttk.Combobox(self, values=sep_methods)
        self.combo_gsm.grid(row=11, column=1, padx=1, pady=1)
        self.combo_gsm.bind("<<ComboboxSelected>>", self.__update_lbl_gsm)

        # Entry эффективность сепарации / диаметр штуцера
        self.ent_eff_gs = tk.Entry(self, width=23, relief=tk.GROOVE,
                                   borderwidth=2)
        self.ent_eff_gs.grid(row=12, column=1, padx=1, pady=1)

        # Кнопка "Сохранить"
        self.btn_save = ttk.Button(self, text='Сохранить', state='disabled',
                                   command=on_save)
        # state=['normal'], command=self.btn_click_save)
        self.btn_save.grid(row=13, column=0, padx=1, pady=1)

        # Кнопка "Рассчитать"
        self.btn_calc = ttk.Button(self, text='Рассчитать', state='disabled',
                                   command=on_calc)
        self.btn_calc.grid(row=13, column=1, padx=1, pady=1)

        # Кнопка "Расчёт в System"
        self.btn_calc_sys = ttk.Button(self, text='Расчёт в System',
                                       state='disabled',
                                       command=on_calc_sys)
        self.btn_calc_sys.grid(row=14, column=0, padx=1, pady=1)

        row = 15
        for text in ['Заб. давление, атм:', 'PI:']:
            lbl = tk.Label(self, text=text, font=FONT_LABELS)
            lbl.grid(row=row, column=0, padx=1, pady=1, sticky='e')
            row += 1

        # Забойное давление текст
        self.lbl_DP_text = tk.Label(self, text='', font=FONT_TEXT)
        self.lbl_DP_text.grid(row=15, column=1, padx=1, pady=1)

        # PI текст
        self.lbl_PI_text = tk.Label(self, text='', font=FONT_TEXT)
        self.lbl_PI_text.grid(row=16, column=1, padx=1, pady=1)

        # Кнопка "Выход"
        self.btn_exit = ttk.Button(self, text='Выход', command=self.destroy)
        self.btn_exit.grid(row=17, column=1, padx=1, pady=1)

    def get_data(self) -> GetDataSchema:
        return GetDataSchema(
            pump_wear=float(self.ent_pump_wear.get().replace(',', '.')),
            correlation=self.combo_corr.get(),
            coef_correlation=float(self.ent_coef_corr.get().replace(',', '.')),
            wht=float(self.ent_whp.get().replace(',', '.')),
            sep_method=self.combo_gsm.get(),
            sep_method_value=float(self.ent_eff_gs.get().replace(',', '.'))
        )

    def set_data(
            self, well_type: str, mrp: int, correlation: str,
            coef_correlation: float, reservoir_pressure: float,
            wht: float, sep_method: str, sep_method_value: float,
            unit: str, pump_wear: float = 0, **kwargs
    ):
        self.lbl_type_text['text'] = well_type
        self.lbl_mrp_text['text'] = mrp
        self.lbl_mrp_text['bg'] = 'red' if mrp < 35 else 'SystemButtonFace'

        self.combo_corr.set(correlation)
        self.ent_coef_corr.insert(0, str(round(coef_correlation, 2)))
        self.lbl_reser_pres_text['text'] = reservoir_pressure
        self.ent_whp.insert(0, str(wht))
        self.ent_pump_wear.insert(0, str(round(pump_wear, 3)))
        self.combo_gsm.set(sep_method)
        self.lbl_gsm['text'] = unit
        self.ent_eff_gs.insert(0, str(sep_method_value))

    def get_field_well(self) -> tuple[str, str]:
        field = self.combo_field.get()
        well = self.combo_well.get()
        return field, well

    def disable_entry_pump_wear(self):
        self.ent_pump_wear['state'] = 'disabled'

    def enable_entry_pump_wear(self):
        self.ent_pump_wear['state'] = 'normal'

    def set_result(
            self, dhp: Optional[float] = None, pi: Optional[float] = None
    ):
        if dhp:
            self.lbl_DP_text['text'] = round(dhp, 1)
        if pi:
            self.lbl_PI_text['text'] = round(pi, 5)
        self.btn_save['state'] = ['normal']

    def reset(self):
        for lbl in [self.lbl_type_text, self.lbl_mrp_text,
                    self.lbl_reser_pres_text,
                    self.lbl_DP_text, self.lbl_PI_text]:
            lbl['text'] = ''

        for ent in [self.ent_coef_corr, self.ent_pump_wear,
                    self.ent_whp, self.ent_eff_gs]:
            ent.delete(0, tk.END)

        self.btn_calc['state'] = ['normal']
        self.btn_calc_sys['state'] = ['normal']
        self.btn_save['state'] = ['disabled']

    def __update_combo_well(self, *args, **kwargs):
        field = self.combo_field.get()
        self.combo_well['values'] = list(self.fields_data[field])
        self.combo_well.delete(0, tk.END)

    def __update_lbl_gsm(self, *args, **kwargs):
        self.ent_eff_gs.delete(0, tk.END)
