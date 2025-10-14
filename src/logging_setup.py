import logging
from typing import Dict, List, Tuple
from datetime import datetime
from logging import Logger


class ListLogHandler(logging.Handler):
    """Handler die logberichten in een lijst opslaat."""
    
    def __init__(self, sink: List[Dict[str, str]]):
        super().__init__()
        self._sink = sink

    def emit(self, record: logging.LogRecord) -> None:
        """Verwerk een log record en voeg toe aan de sink lijst."""
        try:
            ts = datetime.fromtimestamp(record.created).astimezone().strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            ts = ""
        
        msg = record.getMessage()
        level = record.levelname
        self._sink.append({"Tijd": ts, "Niveau": level, "Bericht": msg})


def _build_console_handler() -> logging.Handler:
    """Maak en configureer een console handler."""
    base_fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(base_fmt)
    return console


def _build_memory_handler(sink: List[Dict[str, str]]) -> logging.Handler:
    """Maak en configureer een geheugen handler."""
    memory = ListLogHandler(sink)
    memory.setLevel(logging.INFO)
    memory.setFormatter(logging.Formatter(fmt="%(message)s"))
    return memory


def _reset_logger(logger: Logger) -> None:
    """Reset logger door alle bestaande handlers te verwijderen."""
    if logger.handlers:
        for handler in list(logger.handlers):
            logger.removeHandler(handler)
            try:
                handler.close()
            except Exception:
                pass
    # Voorkom dubbele logging via root logger
    logger.propagate = False


def setup_logging() -> Tuple[Logger, List[Dict[str, str]]]:
    """Initialiseer logging naar console en geheugenlijst."""
    logger = logging.getLogger("validation")
    
    # Reset bestaande handlers
    _reset_logger(logger)
    
    # Configureer logger niveau
    logger.setLevel(logging.INFO)
    
    # Maak sink en handlers
    in_memory_rows: List[Dict[str, str]] = []
    console = _build_console_handler()
    memory = _build_memory_handler(in_memory_rows)
    
    # Voeg handlers toe
    logger.addHandler(console)
    logger.addHandler(memory)
    
    return logger, in_memory_rows
