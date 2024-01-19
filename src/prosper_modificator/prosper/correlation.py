class Correlations:
    data = {
        'Petroleum Experts': 6,
        'Petroleum Experts 2': 7
    }

    @classmethod
    def get_index_by_name(cls, name: str) -> int:
        return cls.data[name]

    @classmethod
    def get_name_by_index(cls, index: int) -> str:
        for name, i in cls.data.items():
            if i == index:
                return name
        else:
            raise ValueError('Некорректный индекс корреляции')

    @classmethod
    def get_correlations(cls) -> list[str]:
        return list(cls.data.keys())
