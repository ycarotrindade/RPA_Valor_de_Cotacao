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

