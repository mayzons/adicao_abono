from datetime import datetime
import sys
import os
import logging
from logging.handlers import TimedRotatingFileHandler


def get_app_and_settings_full_path():
    if getattr(sys, 'frozen', False):
        BASE_PATH = os.path.dirname(sys.executable)
    else:
        BASE_PATH = os.path.dirname(os.path.abspath(__file__))
    return BASE_PATH, os.path.join(BASE_PATH, "Config.ini")


CAM_LOGS_LOGS, CAM_CONFIG_PARSER = get_app_and_settings_full_path()

# Variaveis
hora_atual = datetime.now()
nm_log_data = datetime.strftime(datetime.now(), '%Y-%m-%d')
CAMINHO_LOGS = f'{CAM_LOGS_LOGS}\\LOGS'

nome_log = f'{nm_log_data}'

os.makedirs(CAMINHO_LOGS, exist_ok=True)


logger = logging.getLogger()
cwd = os.getcwd()
handler = TimedRotatingFileHandler(
    f'{CAMINHO_LOGS}\\{nm_log_data}.log',
    when='midnight',
    interval=1,
    backupCount=15)
handler.setFormatter(
    logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s',
                      datefmt='%d/%m/%Y %H:%M:%S'))
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)


def log_warning(message):
    logger.warning(message)


def log_info(message):
    logger.info(message)


def log_debug(message):
    logger.debug(message)


def log_error(message):
    logger.info(message)
