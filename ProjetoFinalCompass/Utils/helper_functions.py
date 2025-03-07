import pandas as pd

def get_jadlog_value(jadlog_service:str):
    JADLOG_DICT = {
        'JADLOG Expresso':"0",
        'JADLOG Econômico':"5",
        'JADLOG Doc':"6",
        'JADLOG Cargo':'12',
        'JADLOG Rodo':'4',
        'JADLOG Package':'3',
        'JADLOG .Com':'9'
    }
    return JADLOG_DICT[jadlog_service]

def calc_finish_task(df_output:pd.DataFrame):
    df_output = df_output[['VALOR COTAÇÃO JADLOG','VALOR COTAÇÃO CORREIOS']]
    total_tasks = len(df_output)
    total_finished = len(df_output.dropna(how='all'))
    total_errors = total_tasks - total_finished
    return total_tasks, total_finished, total_errors