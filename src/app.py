"""QuoteFlow — entry point."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from core.logger import get_logger
from ui.app_core import App

if __name__ == "__main__":
    log = get_logger(__name__)
    log.info("QuoteFlow 啟動")
    App().mainloop()
    log.info("QuoteFlow 結束")
    