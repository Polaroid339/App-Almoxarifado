import logging
from colorama import init, Fore, Style

init()

class ColorFormatter(logging.Formatter):
    FORMATS = {
        logging.DEBUG: Fore.BLUE + "%(message)s" + Style.RESET_ALL,
        logging.INFO: Fore.GREEN + "%(message)s" + Style.RESET_ALL,
        logging.WARNING: Fore.YELLOW + "%(message)s" + Style.RESET_ALL,
        logging.ERROR: Fore.RED + "%(message)s" + Style.RESET_ALL,
        logging.CRITICAL: Style.BRIGHT + Fore.RED + "%(message)s" + Style.RESET_ALL,
    }

    def format(self, record):
        log_fmt = self.FORMATS.get(record.levelno)
        formatter = logging.Formatter(log_fmt)
        return formatter.format(record)

logger = logging.getLogger()
handler = logging.StreamHandler()
handler.setFormatter(ColorFormatter())

logger.addHandler(handler)
logger.setLevel(logging.DEBUG)