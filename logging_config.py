# -*- coding: utf-8 -*-
import logging
import logging.config
from pythonjsonlogger import jsonlogger

LOGGING = {
    "version": 1,
    "disable_existing_loggers": False,
    "formatters": {
        "json": {
            "format": "%(asctime)s - %(levelname)s - %(message)s - %(module)s",
            "class": "pythonjsonlogger.jsonlogger.JsonFormatter",
        }
    },
    "handlers": {
        "stdout": {
            "class": "logging.StreamHandler",
            "stream": "ext://sys.stdout",
            "formatter": "json",
        },
        
        "file": {
            "class": "logging.handlers.RotatingFileHandler",
            "filename": "logging\\logconfig.log",
            "maxBytes": 1024,
            "backupCount": 3,
        }
    },
    "loggers": {"": {"handlers": ["stdout"], "level": "DEBUG"}},
    "loggers": {"": {"handlers": ["file"], "level": "DEBUG"}}
}

logging.config.dictConfig(LOGGING)