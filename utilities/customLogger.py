import logging
import os
from pathlib import Path


class LogGen:
    @staticmethod
    def loggen():
        # Create log file if not exists
        dirpath = os.getcwd()+'\\Logs'
        isExist = os.path.exists(dirpath)
        if isExist:
            filepath = Path(os.getcwd() + "\\Logs\\testlog.log")
            filepath.touch(exist_ok=True)
        elif not isExist:
            os.makedirs(dirpath)
            filepath = Path(os.getcwd() + "\\Logs\\testlog.log")
            filepath.touch(exist_ok=True)
        
        logging.basicConfig(
            filename=".\\Logs\\testlog.log",
            format='%(asctime)s: %(levelname)s: %(message)s',
            datefmt='%m/%d/%Y %I:%M:%S %p',
            force=True
        )
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        return logger
