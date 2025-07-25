from time import sleep

import pandas as pd  # type: ignore
import pyautogui as py  # type: ignore
import pygetwindow as gw  # type: ignore

from utils.config_logger import setup_logger  # type: ignore

logger = setup_logger()

class SAPAutomation:
    def __init__(self, delay: float = 0.8):
        self.delay = delay
        py.FAILSAFE = True  # Enable PyAutoGUI failsafe

    def _safe_click(self, x: int, y: int, clicks: int = 1):
        """Helper to safely move and click."""
        try:
            py.moveTo(x=x, y=y)
            py.click(clicks=clicks)
            sleep(self.delay)
        except py.FailSafeException:
            logger.critical("PyAutoGUI Failsafe triggered. Mouse moved to a corner. Exiting.")
            raise

    def focus_sap(self):
        try:
            window = [w for w in gw.getWindowsWithTitle("SAP Business One") if w.visible][0]
            window.activate()
            window.maximize()
            sleep(1)  
        except IndexError:
            logger.error("Janela do SAP B1 não encontrada.") 
   
    def close_windows_open(self):
        for _ in range(6):
            py.hotkey('ctrl', 'f4')
            sleep(0.3)

    def navigate_to_query_screen(self):
        """Navigates to the SAP query input screen."""
        logger.info("Navigating to query screen...")
       
        corden = py.locateCenterOnScreen(r'C:\Users\g503294\Documents\projetos\contas-pagar\data\images\image.png',confidence=0.8)
        self._safe_click(corden.x, corden.y, clicks=2) # type: ignore
        self._safe_click(x=99, y=334, clicks=2) 
        self._safe_click(x=210, y=539, clicks=2)
        logger.info("Navigation complete.")

    def input_date_parameters(self, start_date: str, end_date: str, current_date: str):
        """
        Inputs date parameters into the SAP application.

        Args:
            start_date (str): The start date (e.g., 'DD/MM/YYYY').
            end_date (str): The end date (e.g., 'DD/MM/YYYY').
            current_date (str): The current date for final input.
        """
        logger.info(f"Inputting date parameters: Start={start_date}, End={end_date}, Current={current_date}")
        py.typewrite(start_date, interval=0.1)
        py.hotkey('tab')
        py.typewrite(end_date, interval=0.1)
        py.hotkey('tab')
        py.hotkey('tab')
        py.hotkey('esc')
        sleep(self.delay)
        py.typewrite(current_date, interval=0.1)
        py.hotkey('enter')
        sleep(8) 
        logger.info("Date parameters input complete.")

    def export_data_to_clipboard(self):
        """Exports data from SAP to the clipboard."""
        logger.info("Exporting data to clipboard...")
        py.rightClick()
        sleep(self.delay)
        py.press('down')
        py.press('down')
        py.press('enter')
        sleep(self.delay)
        logger.info("Data exported to clipboard.")
        return pd.read_clipboard()

