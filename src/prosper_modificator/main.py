from src.prosper_modificator import ProsperModificator

path_res_file = 'result_file.xlsx'

if __name__ == '__main__':
    pm = ProsperModificator(
        path_result_file=path_res_file,
        prosper_files_directory='folder',
        path_init_data='init_data.json',
        is_test=True
    )
