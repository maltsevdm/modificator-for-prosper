def get_excel_column_letter(col_num):
    letter = ''
    while col_num > 0:
        col_num, r = divmod(col_num - 1, 26)
        letter = chr(r + ord('A')) + letter
    return letter


class Column:
    def __init__(self, num):
        self.num = num
        self.letter = get_excel_column_letter(self.num)


class ResultFileCols:
    date: Column = Column(2)
    q_liq_moment: Column = Column(3)
    q_liq: Column = Column(4)
    whp: Column = Column(5)
    q_gaz: Column = Column(6)
    gor: Column = Column(7)
    wc: Column = Column(8)
    frequency: Column = Column(9)
    downtime: Column = Column(10)
    pip: Column = Column(11)
    vlp_ipr_pip: Column = Column(12)
    vlp_ipr_d_pip: Column = Column(13)
    sys_q_liq: Column = Column(14)
    sys_d_q_liq: Column = Column(15)
    sys_pip: Column = Column(16)

    @classmethod
    def get_col_names(cls):
        return list(vars(ResultFileCols)['__annotations__'].keys())

    @classmethod
    def get_points_col_names(cls) -> dict:
        result = {}
        for col_name in vars(cls)['__annotations__'].keys():
            if not (col_name.startswith('vlp') or col_name.startswith('sys')):
                result[col_name] = getattr(cls, col_name).num
        return result
