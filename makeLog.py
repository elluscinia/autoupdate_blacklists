'''
Модуль для настройки работы логирования
'''
import logging.config
import os

import yaml

from resources import resource_path


def setup_logging(default_path=resource_path('logging.yml'),
                  default_level=logging.INFO):
    path = default_path
    if os.path.exists(path):
        with open(path, 'rt') as f:
            config = yaml.safe_load(f.read())
        logging.config.dictConfig(config)
    else:
        logging.basicConfig(level=default_level)

    return logging.getLogger('logger')

logger = setup_logging()
logging.getLogger('requests').setLevel(level=logging.WARNING)
