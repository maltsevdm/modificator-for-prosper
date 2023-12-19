import numpy as np
from openserver import OpenServer
import json
import tkinter as tk
import tkinter.messagebox as mb
from tkinter import ttk
import xlwings as xw
from xlwings.constants import DeleteShiftDirection

from file_paths import path_input_data, path_index_corr, folder, path_WOP, \
    path_prpr_set
from file_settings import COLS_WOP, COLS_PR_SET
from fonts import FONT_LABELS, FONT_TEXT

save_prosper_file = True
num_days = 30


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
    def __init__(self):
        super().__init__()
        self.title('Настройка Prosper')
        self.attributes("-topmost", True)

        # Label Месторождение
        lbl_field = tk.Label(self, text='Месторождение:', font=FONT_LABELS)
        lbl_field.grid(row=1, column=0, padx=1, pady=1, sticky='e')

        # ComboxBox месторождений
        self.combo_field = ttk.Combobox(self, values=list(input_data.keys()))
        self.combo_field.grid(row=1, column=1, padx=1, pady=1)
        self.combo_field.bind("<<ComboboxSelected>>", self.update_combo_well)

        # Label Скважина
        lbl_well = tk.Label(self, text='Скважина:', font=FONT_LABELS)
        lbl_well.grid(row=2, column=0, padx=1, pady=1, sticky='e')

        # ComboxBox скважин
        self.combo_well = ttk.Combobox(self, values=[])
        self.combo_well.grid(row=2, column=1, padx=1, pady=1)

        # Кнопка "Выгрузить"
        btn_load = ttk.Button(self, text='Выгрузить',
                              command=self.btn_click_load)
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
        self.ent_temp_buf = tk.Entry(self, width=23, relief=tk.GROOVE,
                                     borderwidth=2)
        self.ent_temp_buf.grid(row=10, column=1, padx=1, pady=1)

        # Combobox методов сепарации газа
        self.combo_gsm = ttk.Combobox(self, values=gsm_list)
        self.combo_gsm.grid(row=11, column=1, padx=1, pady=1)
        self.combo_gsm.bind("<<ComboboxSelected>>", self.update_lbl_gsm)

        # Entry эффективность сепарации / диаметр штуцера
        self.ent_eff_gs = tk.Entry(self, width=23, relief=tk.GROOVE,
                                   borderwidth=2)
        self.ent_eff_gs.grid(row=12, column=1, padx=1, pady=1)

        # Кнопка "Сохранить"
        self.btn_save = ttk.Button(self, text='Сохранить', state='disabled',
                                   command=self.btn_click_save)
        # state=['normal'], command=self.btn_click_save)
        self.btn_save.grid(row=13, column=0, padx=1, pady=1)

        # Кнопка "Рассчитать"
        self.btn_calc = ttk.Button(self, text='Рассчитать', state='disabled',
                                   command=self.btn_click_calc)
        self.btn_calc.grid(row=13, column=1, padx=1, pady=1)

        # Кнопка "Расчёт в System"
        self.btn_calc_sys = ttk.Button(self, text='Расчёт в System',
                                       state='disabled',
                                       command=self.btn_click_calc_sys)
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

    def btn_click_load(self):
        global fountain, TMS_work, ws_well

        for lbl in [self.lbl_type_text, self.lbl_mrp_text,
                    self.lbl_reser_pres_text,
                    self.lbl_DP_text, self.lbl_PI_text]:
            lbl['text'] = ''
        for ent in [self.ent_coef_corr, self.ent_pump_wear,
                    self.ent_temp_buf, self.ent_eff_gs]:
            ent.delete(0, tk.END)
        self.btn_calc['state'] = ['normal']
        self.btn_calc_sys['state'] = ['normal']
        self.btn_save['state'] = ['disabled']
        ws_pr_set.cells(1, 10).value = None
        ws_pr_set.cells(1, 13).value = None

        field = get_value_from_cb(self.combo_field, 'Месторождение', input_data)
        if field is None: return
        well = get_value_from_cb(self.combo_well, 'Скважина', input_data[field])
        if well is None: return

        print('Выгружаю...')
        wb_Prosper_setting.app.screen_updating = False

        sheetname = f'{well}_{field.upper()[0]}'
        ws_well = wb_wop.sheets[sheetname]

        fountain = type_well(ws_well)  # Фонтан или нет
        if fountain:
            self.lbl_type_text['text'] = 'Фонтан'
            self.ent_pump_wear['state'] = 'disabled'
        else:
            self.ent_pump_wear['state'] = 'normal'
            TMS_work = check_tms(ws_well)
            self.lbl_type_text['text'] = ('ЭЦН с датчиком ТМС' if TMS_work
                                          else 'ЭЦН без датчика ТМС')

        # Очистка листа
        while ws_pr_set.cells(5, 3).value:
            ws_pr_set.range('5:5').api.Delete(DeleteShiftDirection.xlShiftUp)

        self.add_mrp(ws_well)
        open_prosper_file(field, well)

        # Получение корреляции
        correlation = c.DoGet('PROSPER.ANL.SYS.TubingLabel')
        self.combo_corr.current(correlations.index(correlation))

        # Получение коэффициента корреляции
        self.ent_coef_corr.insert(
            0,
            round(c.DoGet(f'PROSPER.ANL.COR.Corr[{{{correlation}}}].A[0]'), 2)
        )

        # Пластовое давление
        self.lbl_reser_pres_text['text'] = input_data[field][well][
            'reservoir_pres']

        # Температура на буфере
        self.ent_temp_buf.insert(0, '10')

        if not fountain:
            # Получение износа насоса
            self.ent_pump_wear.insert(
                0, round(c.DoGet('PROSPER.SIN.ESP.Wear'), 3))

        gsm = c.DoGet('PROSPER.SIN.ESP.SepMethod')
        self.combo_gsm.set(gsm_list[gsm])
        if gsm == 0:
            unit = c.DoGet('PROSPER.SIN.ESP.Efficiency.Unitname')
            self.lbl_gsm[
                'text'] = f'Эффективность, {"%:" if unit == "percent" else "д.е.:"}'
            self.ent_eff_gs.insert(0, c.DoGet('PROSPER.SIN.ESP.Efficiency'))
        elif gsm == 1:
            unit = c.DoGet('PROSPER.SIN.ESP.SepPortDiam.Unitname')
            self.lbl_gsm['text'] = f'Диаметр, {unit}'
            self.ent_eff_gs.insert(0,
                                   round(c.DoGet('PROSPER.SIN.ESP.SepPortDiam'),
                                         3))

        # Буквы колонок
        let_q_liq = excel_column_letter(COLS_PR_SET['inp_data']['Q_liq'])
        let_q_gaz = excel_column_letter(COLS_PR_SET['inp_data']['Q_gaz'])
        let_wc = excel_column_letter(COLS_PR_SET['inp_data']['WC'])
        let_downtime = excel_column_letter(COLS_PR_SET['inp_data']['downtime'])

        row_pr_set = 5
        for row_da in range(12 + num_days, 11, -1):
            if ws_well.cells(row_da, COLS_WOP['Q_liq']).value in [None, '-']:
                continue

            for param, col in COLS_PR_SET['inp_data'].items():
                if param in ['GF', 'Q_liq_moment']:
                    continue
                elif param in ['WC', 'Q_gaz']:
                    row_last_val = row_da
                    while (ws_well.cells(row_last_val,
                                         COLS_WOP[param]).value in [None, '-']
                           and ws_well.cells(row_last_val, COLS_WOP[
                                'date']).value is not None):
                        row_last_val += 1
                    ws_pr_set.cells(row_pr_set, col).value = ws_well.cells(
                        row_last_val, COLS_WOP[param]).value
                else:
                    ws_pr_set.cells(row_pr_set, col).value = ws_well.cells(
                        row_da, COLS_WOP[param]).value

            let_q_liq_mom = excel_column_letter(
                COLS_PR_SET['inp_data']['Q_liq_moment'])

            # Формула для дебита жидкости мгновенного
            ws_pr_set.cells(row_pr_set,
                            COLS_PR_SET['inp_data']['Q_liq_moment']).value = (
                f'={let_q_liq}{row_pr_set}/(1-{let_downtime}{row_pr_set})')

            # Формула для ГФ
            ws_pr_set.cells(row_pr_set, COLS_PR_SET['inp_data']['GF']).value = (
                f'=IF({let_q_gaz}{row_pr_set}/({let_q_liq_mom}{row_pr_set}'
                f'*(100-{let_wc}{row_pr_set})/100)<256, '
                f'256, '
                f'{let_q_gaz}{row_pr_set}/({let_q_liq_mom}{row_pr_set}'
                f'*(100-{let_wc}{row_pr_set})/100))')

            ws_pr_set.range(
                f'{row_pr_set}:{row_pr_set}').api.HorizontalAlignment = (
                xw.constants.HAlign.xlHAlignCenter)

            row_pr_set += 1

        wb_Prosper_setting.app.screen_updating = True
        print('Данные выгружены.')

    def btn_click_calc(self):
        field = get_value_from_cb(self.combo_field, 'Месторождение', input_data)
        if field is None:
            return
        well = get_value_from_cb(self.combo_well, 'Скважина', input_data[field])
        if well is None:
            return
        correlation = get_value_from_cb(self.combo_corr, 'Корреляция',
                                        correlations)
        if not correlation is None:
            return

        if not fountain:
            pump_wear = get_value_from_ent(self.ent_pump_wear,
                                           'коэф. износа насоса')
            if pump_wear is None:
                return
            gsm = get_value_from_cb(self.combo_gsm, 'Метод сепарации', gsm_list)
            if gsm is None:
                return
            gsm_index = gsm_list.index(gsm)
            gsm_val = get_value_from_ent(self.ent_eff_gs, gsm)
            if gsm_val is None:
                return
        coef_corr = get_value_from_ent(self.ent_coef_corr, 'коэф. корреляции')
        if coef_corr is None:
            return
        temp_buf = get_value_from_ent(self.ent_temp_buf, 'темп. на буфере')
        if temp_buf is None:
            return

        green_deltas = 0
        yellow_deltas = 0
        deltas = []

        open_prosper_file(field, well)
        print('Рассчитываю...')

        # Запись коэффициента корреляции
        c.DoSet(f'PROSPER.ANL.COR.Corr[{{{correlation}}}].A[0]', coef_corr)

        if not fountain:
            # Запись метода сепарации и значения
            choise_gsm(gsm_index, gsm_val)
            # Запись износа насоса
            c.DoSet('PROSPER.SIN.ESP.Wear', pump_wear)

        if fountain or not TMS_work:
            if not fountain:
                # Запись частоты
                c.DoSet('PROSPER.ANL.WHP.Data[0].Freq',
                        input_data[field][well]['frequency'])
                c.DoSet('PROSPER.SIN.ESP.Frequency',
                        input_data[field][well]['frequency'])

            gor_ds = input_data[field][well]['gor']
            wc_ds = input_data[field][well]['wc']
            q_liq_ds = input_data[field][well]['q_liq']
            pres_buf_ds = input_data[field][well]['pres_buf']

            # Забойное давление
            dhp = calc_dhp_in_bhp_from_whp(
                ws_pr_set.cells(5, COLS_PR_SET['inp_data']['date']).value,
                q_liq_ds, pres_buf_ds, temp_buf, gor_ds, wc_ds, correlation)
            self.lbl_DP_text['text'] = round(dhp, 1)

            reser_pres = input_data[field][well]['reservoir_pres']

            if dhp >= reser_pres:
                mb.showerror('Предупреждение',
                             'Расчётное забойное давление больше пластового!')
                return

            pi = calc_pi(reser_pres, wc_ds, gor_ds, q_liq_ds, dhp)
            self.lbl_PI_text['text'] = round(pi, 5)

            for row_pr_set in range(5, 50):
                if ws_pr_set.cells(
                        row_pr_set,
                        COLS_PR_SET['inp_data']['date']).value is None:
                    break

                system_calc(
                    ws_pr_set.cells(row_pr_set,
                                    COLS_PR_SET['inp_data']['pres_buf']).value,
                    ws_pr_set.cells(row_pr_set,
                                    COLS_PR_SET['inp_data']['WC']).value,
                    ws_pr_set.cells(row_pr_set,
                                    COLS_PR_SET['inp_data']['GF']).value,
                    correlation)

                for pr_param, param in {'QLIQ': 'Q_liq', 'PIP': 'Ppr'}.items():
                    if fountain and param == 'Ppr':
                        continue

                    # Получение дебита жидкости из System
                    ws_pr_set.cells(row_pr_set,
                                    COLS_PR_SET['sys'][param]).value = round(
                        c.DoGet(f'PROSPER.OUT.SYS.SOL[0].{pr_param}'), 1)

                    if param != 'Q_liq':
                        continue

                    # Формула для расчёта дельты
                    ws_pr_set.cells(row_pr_set,
                                    COLS_PR_SET['sys']['d_' + param]).value = (
                        formula_d.format(
                            let_1=excel_column_letter(
                                COLS_PR_SET['sys'][param]),
                            let_2=excel_column_letter(
                                COLS_PR_SET['inp_data']['Q_liq_moment']),
                            row=row_pr_set))

                    # Покраска ячейки с дельтой
                    fill_delta_cell(
                        ws_pr_set.cells(row_pr_set,
                                        COLS_PR_SET['sys']['d_' + param]))

        else:
            row_num = c.DoGet('PROSPER.ANL.VMT.Data.Count')
            for row_pr_set in range(5, 50):
                if ws_pr_set.cells(
                        row_pr_set,
                        COLS_PR_SET['inp_data']['date']).value is None:
                    break

                frequency = ws_pr_set.cells(
                    row_pr_set, COLS_PR_SET['inp_data']['frequency']).value
                whp = ws_pr_set.cells(
                    row_pr_set, COLS_PR_SET['inp_data']['pres_buf']).value
                wc = ws_pr_set.cells(
                    row_pr_set, COLS_PR_SET['inp_data']['WC']).value
                q_liq = ws_pr_set.cells(
                    row_pr_set, COLS_PR_SET['inp_data']['Q_liq_moment']).value
                gor = ws_pr_set.cells(
                    row_pr_set, COLS_PR_SET['inp_data']['GF']).value
                # Занесение частоты в ESP Input Data
                c.DoSet('PROSPER.SIN.ESP.Frequency', frequency)

                date_format = ws_pr_set.cells(
                    row_pr_set, COLS_PR_SET['inp_data']['date']).value.replace(
                    '.', '/')

                # Запись данных в VLP/IPR
                add_data_in_VLP_IPR(
                    row_num, date_format, whp, temp_buf, wc, q_liq,
                    ws_pr_set.cells(row_pr_set,
                                    COLS_PR_SET['inp_data']['Ppr']).value,
                    float(self.lbl_reser_pres_text['text']),
                    gor, frequency, pump_wear)
                row_num += 1

                # Получение давления на приёме
                pip = calc_pres_intake_in_tcc(
                    whp, wc, gor, q_liq, correlation,
                    ws_pr_set.cells(
                        row_pr_set, COLS_PR_SET['inp_data']['Ppr']).value
                )

                ws_pr_set.cells(
                    row_pr_set, COLS_PR_SET['vlp_ipr']['Ppr']
                ).value = (round(pip, 1))

                if pip == -1:
                    ws_pr_set.cells(
                        row_pr_set, COLS_PR_SET['vlp_ipr']['d_Ppr']
                    ).value = None
                    ws_pr_set.cells(
                        row_pr_set, COLS_PR_SET['vlp_ipr']['d_Ppr']
                    ).color = (255, 0, 0)
                    continue

                # Формула для расчёта дельты
                ws_pr_set.cells(row_pr_set,
                                COLS_PR_SET['vlp_ipr']['d_Ppr']).value = (
                    formula_d.format(
                        let_1=excel_column_letter(
                            COLS_PR_SET['vlp_ipr']['Ppr']),
                        let_2=excel_column_letter(
                            COLS_PR_SET['inp_data']['Ppr']),
                        row=row_pr_set))

                # Покраска ячейки с дельтой
                fill_delta_cell(
                    ws_pr_set.cells(row_pr_set,
                                    COLS_PR_SET['vlp_ipr']['d_Ppr']))

                delta_Ppr = ws_pr_set.cells(row_pr_set, COLS_PR_SET['vlp_ipr'][
                    'd_Ppr']).value
                deltas.append(delta_Ppr)

                if delta_Ppr <= 5:
                    green_deltas += 1
                elif delta_Ppr <= 10:
                    yellow_deltas += 1

            set_rate = (green_deltas + yellow_deltas * 0.5) / len(deltas)
            ws_pr_set.cells(1, 10).value = round(set_rate, 2)
            ws_pr_set.cells(1, 13).value = round(np.median(deltas), 1)
            print('Среднеквадратичное отклонение дельт: ', np.std(deltas))
        self.btn_save['state'] = ['normal']
        mb.showinfo('Готово', 'Готово!')
        print('Расчёт закончен.')

    def add_mrp(self, ws):
        try:
            mrp = int(ws.cells(5, 3).value)
            self.lbl_mrp_text['text'] = mrp
            self.lbl_mrp_text['bg'] = 'red' if mrp < 35 else 'SystemButtonFace'
        except TypeError:
            self.lbl_mrp_text['text'] = None

    def btn_click_save(self):
        field = get_value_from_cb(self.combo_field, 'Месторождение', input_data)
        if field is None: return
        well = get_value_from_cb(self.combo_well, 'Скважина', input_data[field])
        if well is None: return
        if not fountain:
            pump_wear = get_value_from_ent(self.ent_pump_wear,
                                           'коэф. износа насоса')
            if pump_wear is None: return
            gsm = get_value_from_cb(self.combo_gsm, 'Метод сепарации', gsm_list)
            if gsm is None: return
            gsm_index = gsm_list.index(gsm)
            gsm_val = get_value_from_ent(self.ent_eff_gs, gsm)
            if gsm_val is None: return
        coef_corr = get_value_from_ent(self.ent_coef_corr, 'коэф. корреляции')
        if coef_corr is None: return
        correlation = get_value_from_cb(self.combo_corr, 'Корреляция',
                                        correlations)
        if correlation is None: return
        temp_buf = get_value_from_ent(self.ent_temp_buf, 'темп. на буфере')
        if temp_buf is None: return

        print('Сохраняю...')

        if not fountain and TMS_work:
            row_save = None
            ask_form = AskRowForm(self)
            row_save = ask_form.open()

            open_prosper_file(field, well)

            # Запись метода сепарации и значения
            choise_gsm(gsm_index, gsm_val)

            # Параметры из Datasheet
            reser_pres = input_data[field][well]['reservoir_pres']
            pres_buf_ds = input_data[field][well]['pres_buf']
            wc_ds = input_data[field][well]['wc']
            gor_ds = input_data[field][well]['gor']
            q_liq_ds = input_data[field][well]['q_liq']
            freq_ds = input_data[field][well]['frequency']
            pip_ds = input_data[field][well]['pip']

            # Коэффициент корреляции
            c.DoSet(f'PROSPER.ANL.COR.Corr[{{{correlation}}}].A[0]', coef_corr)
            # Износ насоса
            c.DoSet('PROSPER.SIN.ESP.Wear', pump_wear)
            # Запись частоты
            c.DoSet('PROSPER.SIN.ESP.Frequency', freq_ds)

            date_format = ws_pr_set.cells(
                row_save, COLS_PR_SET['inp_data']['date']).value.replace('.',
                                                                         '/')

            row_num = c.DoGet('PROSPER.ANL.VMT.Data.Count')
            # Запись данных в VLP/IPR
            add_data_in_VLP_IPR(
                row_num, date_format,
                ws_pr_set.cells(row_save,
                                COLS_PR_SET['inp_data']['pres_buf']).value,
                temp_buf,
                ws_pr_set.cells(row_save, COLS_PR_SET['inp_data']['WC']).value,
                ws_pr_set.cells(row_save,
                                COLS_PR_SET['inp_data']['Q_liq_moment']).value,
                ws_pr_set.cells(row_save, COLS_PR_SET['inp_data']['Ppr']).value,
                reser_pres,
                ws_pr_set.cells(row_save, COLS_PR_SET['inp_data']['GF']).value,
                ws_pr_set.cells(row_save,
                                COLS_PR_SET['inp_data']['frequency']).value,
                pump_wear)

            # Получение расчётного забойного давления
            dhp = calc_pres_downhole_in_tcc(
                pres_buf_ds, wc_ds, gor_ds, q_liq_ds, correlation, pip_ds)
            self.lbl_DP_text['text'] = round(dhp, 1)
            if dhp >= reser_pres:
                mb.showerror('Предупреждение',
                             'Расчётное забойное давление больше пластового!')
                return

            pi = calc_pi(reser_pres, wc_ds, gor_ds, q_liq_ds, dhp)
            self.lbl_PI_text['text'] = round(pi, 5)

        # Корреляция в System
        c.DoSet('PROSPER.ANL.SYS.TubingLabel', correlation)
        # Корреляция в VLP
        c.DoSet('PROSPER.ANL.VLP.TubingLabel', correlation)

        if save_prosper_file:
            add_params_to_ws_well(correlation, coef_corr,
                                  pump_wear if not fountain else None)

            path_file_prosper = f'{folder}\\{input_data[field][well]["prosper_fn"]}'
            c.DoCmd(f'PROSPER.SAVEFILE({path_file_prosper})')
        mb.showinfo('Сохранено', 'Файл Prosper сохранён.')
        print('Файл Prosper сохранён.')

    def btn_click_calc_sys(self):
        field = get_value_from_cb(self.combo_field, 'Месторождение', input_data)
        if field is None:
            return
        well = get_value_from_cb(self.combo_well, 'Скважина', input_data[field])
        if well is None:
            return
        correlation = get_value_from_cb(self.combo_corr, 'Корреляция',
                                        correlations)
        if correlation is None:
            return
        if not fountain:
            pump_wear = get_value_from_ent(self.ent_pump_wear,
                                           'коэф. износа насоса')
            if pump_wear is None:
                return
            gsm = get_value_from_cb(self.combo_gsm, 'Метод сепарации', gsm_list)
            if gsm is None:
                return
            gsm_index = gsm_list.index(gsm)
            gsm_val = get_value_from_ent(self.ent_eff_gs, gsm)
            if gsm_val is None:
                coef_corr = get_value_from_ent(self.ent_coef_corr,
                                               'коэф. корреляции')
                if coef_corr is None:
                    return

        open_prosper_file(field, well)
        # Запись коэффициента корреляции
        c.DoSet(f'PROSPER.ANL.COR.Corr[{{{correlation}}}].A[0]', coef_corr)
        if not fountain:
            # Запись метода сепарации и значения
            choise_gsm(gsm_index, gsm_val)
            # Запись износа насоса
            c.DoSet('PROSPER.SIN.ESP.Wear', pump_wear)

        print('Рассчитываю...')
        for row in range(5, 50):
            if ws_pr_set.cells(row,
                               COLS_PR_SET['inp_data']['date']).value is None:
                break

            whp = ws_pr_set.cells(row,
                                  COLS_PR_SET['inp_data']['pres_buf']).value
            wc = ws_pr_set.cells(row, COLS_PR_SET['inp_data']['WC']).value
            gor = ws_pr_set.cells(row, COLS_PR_SET['inp_data']['GF']).value

            system_calc(whp, wc, gor, correlation)

            ws_pr_set.cells(row, COLS_PR_SET['sys']['Q_liq']).value = round(
                c.DoGet('PROSPER.OUT.SYS.SOL[0].QLIQ'), 1)

            ws_pr_set.cells(row, COLS_PR_SET['sys'][
                'd_Q_liq']).value = formula_d.format(
                let_1=excel_column_letter(COLS_PR_SET['sys']['Q_liq']),
                let_2=excel_column_letter(
                    COLS_PR_SET['inp_data']['Q_liq_moment']),
                row=row)
            # Покраска ячейки с дельтой
            fill_delta_cell(ws_pr_set.cells(row, COLS_PR_SET['sys']['d_Q_liq']))

    def update_combo_well(self, event):
        field = self.combo_field.get()
        self.combo_well['values'] = list(input_data[field].keys())
        self.combo_well.delete(0, tk.END)

    def update_lbl_gsm(self, event):
        gsm = self.combo_gsm.get()
        if gsm == gsm_list[0]:
            self.lbl_gsm['text'] = 'Эффективность, %:'
        elif gsm == gsm_list[1]:
            self.lbl_gsm['text'] = 'Диаметр, mm:'
        self.ent_eff_gs.delete(0, tk.END)


def get_value_from_cb(cb, name, check_list):
    item = cb.get()
    if item == '':
        mb.showerror('Ошибка',
                     f'Поле "{name}" не может быть пустым!')
        return None
    elif item not in check_list:
        mb.showerror('Ошибка',
                     f'Неверно указано {name}!')
        return None
    else:
        return item


def get_value_from_ent(ent, name):
    try:
        value = float(ent.get().replace(',', '.'))
        return value
    except (TypeError, ValueError):
        mb.showerror('Ошибка', f'Неверно указан {name}!')
        return None


def excel_column_letter(col_num):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    letter = ''
    while col_num > 0:
        col_num, r = divmod(col_num - 1, 26)
        letter = chr(r + ord('A')) + letter
    return letter


def add_params_to_ws_well(corr, coef_corr, pump_wear):
    for param, var in {'corr': corr, 'coef_corr': coef_corr,
                       'pump_wear': pump_wear}.items():
        cell = ws_well.cells(12, COLS_WOP[param])
        cell.value = var
        cell.color = (255, 242, 204)
        # Горизонтальное выравнивание по центру
        cell.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter


def get_median_value(col) -> float:
    values = []
    for row in range(5, 50):
        if ws_pr_set.cells(row, col).value is None:
            break
        if (col == 8 and ws_pr_set.cells(row, col).value
                != ws_pr_set.cells(row - 1, col).value):
            values.append(ws_pr_set.cells(row, col).value)
        else:
            values.append(ws_pr_set.cells(row, col).value)
    return np.median(values)


def calc_pi(reser_pres, wc, gor, q_liq, pres_downhole):
    ''' Расчёта PI в секции IPR'''
    c.DoSet('PROSPER.SIN.IPR.Single.IprMethod', 1)  # Vogel
    c.DoSet('PROSPER.SIN.IPR.Single.Pres', reser_pres)
    c.DoSet('PROSPER.SIN.IPR.Single.Tres', 104)
    c.DoSet('PROSPER.SIN.IPR.Single.WC', wc)
    c.DoSet('PROSPER.SIN.IPR.Single.totgor', gor)
    c.DoSet('PROSPER.SIN.IPR.Single.Qtest', q_liq)
    c.DoSet('PROSPER.SIN.IPR.Single.Ptest', pres_downhole)
    c.DoCmd("PROSPER.REFRESH")
    c.DoCmd("PROSPER.IPR.Calc")

    pi = c.DoGet('PROSPER.Sin.IPR.Single.PINSAV')

    c.DoSet('PROSPER.SIN.IPR.Single.IprMethod', 0)  # PI Entry
    c.DoSet('PROSPER.SIN.IPR.Single.Pindex', pi)
    c.DoCmd("PROSPER.REFRESH")
    c.DoCmd("PROSPER.IPR.Calc")
    return pi


def calc_pres_intake_in_tcc(pres_buf, wc, gor, q_liq, correlation, pip_fact):
    add_data_in_TCC(pres_buf, wc, gor, q_liq, correlation, pip_fact)

    value_regime = list(c.DoGet(
        f'PROSPER.OUT.TCC.RES[{index_corr["OUT_TCC"][correlation]}][0][$].Regime'))
    try:
        index_pip = value_regime.index('Pump Intake')
    except ValueError:
        ws_pr_set.cells(row_pr_set,
                        COLS_PR_SET['vlp_ipr']['Ppr']).value = 'Error'
        ws_pr_set.cells(row_pr_set,
                        COLS_PR_SET['vlp_ipr']['d_Ppr']).value = None
        print(f'Строка {row_pr_set} | Не найдена точка насоса!')
        return -1
    return c.DoGet(
        f'PROSPER.OUT.TCC.RES[{index_corr["OUT_TCC"][correlation]}][0][{index_pip}].Pres')


def open_prosper_file(field, well):
    print('Открываю файл проспер')
    path_file_prosper = f'{folder}\\{input_data[field][str(well)]["prosper_fn"]}'
    c.DoCmd(f'PROSPER.OPENFILE({path_file_prosper})')
    print('Файл проспера открыт')


def calc_pres_downhole_in_tcc(pres_buf, wc, gor, q_liq, correlation, pip_fact):
    add_data_in_TCC(pres_buf, wc, gor, q_liq, correlation, pip_fact)
    value_regime = list(c.DoGet(
        f'PROSPER.OUT.TCC.RES[{index_corr["OUT_TCC"][correlation]}][0][$].Regime'))
    index_bhp = value_regime.index('Bottomhole')
    return c.DoGet(
        f'PROSPER.OUT.TCC.RES[{index_corr["OUT_TCC"][correlation]}][0][{index_bhp}].Pres')


def add_data_in_VLP_IPR(row, date, pres_buf, temp_buf, WC, Q_liq, Ppr,
                        reser_pres, GOR, freq, pump_wear):
    ''' Запись данных в секцию VLP/IPR '''
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].Date', date)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].THpres', pres_buf)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].THtemp', temp_buf)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].WC', WC)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].Rate', Q_liq)
    pump_depth = c.DoGet('PROSPER.SIN.ESP.Depth')
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].Gdepth', pump_depth)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].Gpres', Ppr)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].Pres', reser_pres)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].GOR', GOR)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].GORfree', 0)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].Freq', freq)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].Wear', pump_wear)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].PIP', Ppr)
    c.DoSet(f'PROSPER.ANL.VMT.Data[{row}].PDP', 2)


def add_data_in_TCC(pres_buf, wc, gor, q_liq, correlation, pip_fact=None):
    ''' Запись данных в секцию TCC '''
    c.DoSet('PROSPER.ANL.TCC.Pres', pres_buf)
    c.DoSet('PROSPER.ANL.TCC.WC', wc)
    c.DoSet('PROSPER.ANL.TCC.GOR', gor)
    c.DoSet('PROSPER.ANL.TCC.GORFree', 0)
    c.DoSet('PROSPER.ANL.TCC.Rate', q_liq)
    for i in range(0, 10):
        if c.DoGet(f'PROSPER.ANL.TCC.Comp[{i}].Msd') is None:
            break
        else:
            c.DoSet(f'PROSPER.ANL.TCC.Comp[{i}].Msd', '')
            c.DoSet(f'PROSPER.ANL.TCC.Comp[{i}].Prs', '')
    c.DoSet('PROSPER.ANL.TCC.Comp[0].Msd', c.DoGet('PROSPER.SIN.ESP.Depth'))
    c.DoSet('PROSPER.ANL.TCC.Comp[0].Prs', pip_fact)
    # Проставление флажка около нужной корреляции в TCC
    c.DoSet('PROSPER.ANL.TCC.CORR [$]', 0)
    c.DoSet(f'PROSPER.ANL.TCC.CORR[{index_corr["OUT_TCC"][correlation]}]', 1)
    c.DoCmd('PROSPER.REFRESH')
    c.DoCmd('PROSPER.ANL.TCC.CALC')


def type_well(ws) -> bool:
    # Проверка фонтан или нет
    for row in range(12 + num_days, 11, -1):
        if ws.cells(row, COLS_WOP['frequency']).value not in ['-', None]:
            return False
    return True


def check_tms(ws) -> bool:
    for row in range(12 + num_days, 11, -1):
        if ws.cells(row, COLS_WOP['Ppr']).value not in ['-', None]:
            return True
    return False


def calc_dhp_in_bhp_from_whp(date, q_liq, whp, wht, gor, wc, corr):
    # Запись в BHP from WHP
    c.DoSet('PROSPER.ANL.WHP.Data[0].Time', date)
    c.DoSet('PROSPER.ANL.WHP.Data[0].Rate', q_liq)
    c.DoSet('PROSPER.ANL.WHP.Data[0].WHP', whp)
    c.DoSet('PROSPER.ANL.WHP.Data[0].WHT', wht)
    c.DoSet('PROSPER.ANL.WHP.Data[0].GASF', gor)
    c.DoSet('PROSPER.ANL.WHP.Data[0].WATF', wc)
    c.DoSet('PROSPER.ANL.WHP.TubingLabel', corr)
    c.DoCmd('PROSPER.REFRESH')
    c.DoCmd('PROSPER.ANL.WHP.CALC')
    return c.DoGet('PROSPER.ANL.WHP.Data[0].BHP')


def system_calc(whp, wc, gor, correlation):
    # Расчёт в System
    c.DoSet('PROSPER.ANL.SYS.Pres', whp)
    c.DoSet('PROSPER.ANL.SYS.WC', wc)
    c.DoSet('PROSPER.ANL.SYS.GOR', gor)
    c.DoSet('PROSPER.ANL.SYS.TubingLabel', correlation)
    c.DoCmd('PROSPER.ANL.SYS.Calc')


def fill_delta_cell(cell, green_max=5, yel_max=10):
    # Покраска ячейки с дельтой
    if abs(cell.value) <= green_max:
        cell.color = (146, 208, 80)
    elif abs(cell.value) <= yel_max:
        cell.color = (255, 255, 0)
    else:
        cell.color = (255, 0, 0)


def choise_gsm(gsm_index, gsm_val):
    c.DoSet('PROSPER.SIN.ESP.SepMethod', gsm_index)
    if gsm_index == 0:
        c.DoSet('PROSPER.SIN.ESP.Efficiency', gsm_val)
    elif gsm_index == 1:
        c.DoSet('PROSPER.SIN.ESP.RgsDPortMethod', 1)
        c.DoSet('PROSPER.SIN.ESP.SepPortDiam', gsm_val)


if __name__ == '__main__':
    with open(path_input_data, encoding='windows-1251') as file:
        input_data = json.load(file)

    with open(path_index_corr, encoding='windows-1251') as file:
        index_corr = json.load(file)
    correlations = list(index_corr['OUT_TCC'].keys())

    wb_Prosper_setting = xw.Book(path_prpr_set)
    ws_pr_set = wb_Prosper_setting.sheets('Настройка')
    wb_wop = xw.Book(path_WOP)
    gsm_list = ['Эффективность сепарации', 'Диаметр штуцера']
    formula_d = '=({let_1}{row}-{let_2}{row})*100/{let_2}{row}'

    fountain = False
    TMS_work = False
    ws_well = None

    c = OpenServer()
    c.connect()
    try:
        app = App()
        app.mainloop()
    except Exception as ex:
        print(ex)
    finally:
        c.disconnect()
