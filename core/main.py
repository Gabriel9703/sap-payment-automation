import pandas as pd  # type: ignore
from pyautogui import FailSafeException  # type: ignore

from core.automation_sap import SAPAutomation
from core.processing_data import ExcelFormater, ExcelProcessor
from utils.config_logger import setup_logger
from utils.config_path import (
    PATH_INVOICES_MERGE,
    PATH_INVOICES_PAID,
    PATH_INVOICES_TO_PAY,
)

logger = setup_logger()

def run_automation_sap(*, start_date: str, end_date: str, current_date: str) -> pd.DataFrame | None:
    """" Roda a automação no SAP com base nas datas fornecidas e retorna um DataFrame com os dados exportados.
    
    Args:
        start_date (str): Data de início para a consulta no formato 'DD/MM/YYYY'.
        end_date (str): Data de fim para a consulta no formato 'DD/MM/YYYY'.
        current_date (str): Data atual para o processamento no formato 'DD/MM/YYYY'.
    Returns:
            pd.DataFrame | None: DataFrame com os dados exportados do SAP ou None em caso de erro.
    Raises:
        Exception: Se houver um erro durenate o processo de gravação no excel.
        FailSafeException: Se o usuário interromper a automação com PyAutoGUI failsafe."""
    sap_automator = SAPAutomation(delay=0.8) 
    
    logger.info("Starting SAP invoice processing automation.")
    try:
        sap_automator.focus_sap()
        sap_automator.close_windows_open()
        sap_automator.navigate_to_query_screen()    
        sap_automator.input_date_parameters(
            start_date = start_date,
            end_date = end_date,
            current_date = current_date
        )
        current_sap_data_df = sap_automator.export_data_to_clipboard()
        return current_sap_data_df

    except FailSafeException:
        logger.critical("Automation terminated due to PyAutoGUI failsafe. User intervention likely.")
    except Exception as e:
        logger.critical(f"An unexpected error occurred: {e}", exc_info=True)
        return None
    
def export_report_sap_to_excel(df_sap_to_export: pd.DataFrame) -> None:
    """Exporta o DataFrame retirado do SAP para um arquivo Excel.
    Args:
        df_to_export (pd.DataFrame): DataFrame contendo os dados a serem exportados.
    Raises:
        Exception: Se houver um erro durante a exportação do DataFrame para o Excel."""
    try:
        excel_processor = ExcelFormater(file_path=PATH_INVOICES_TO_PAY, df_to_export_excel = df_sap_to_export )   
        excel_processor.export_df_to_excel(sheet_name='em_aberto', start_row_ex=0, start_col_ex=0)
    except Exception as e:
        logger.error(f"Failed to export DataFrame to Excel: {e}")
        raise Exception(f"Failed to export DataFrame to Excel: {e}") 
    
def run_processor_excel(target_date: str = '30/07/2025') -> tuple[pd.DataFrame, pd.DataFrame]:
    """Processa os dados de faturas a pagar e pagas, reconciliando e formatando os DataFrames.
    
    Args:
        target_date (str): Data alvo para a consulta no formato 'DD/MM/YYYY'.
    Returns:
        tuple[pd.DataFrame, pd.DataFrame]: Uma tupla contendo os DataFrames pagos e a pagar.
   """
    # 1 - Load existing 'notas_pagas' and process
    invoices_to_pay = ExcelProcessor(file_path=PATH_INVOICES_TO_PAY)
    invoices_paid = ExcelProcessor(file_path=PATH_INVOICES_PAID)   

    # 2 - Load data from Excel files into DataFrames
    df_invoices_paid = invoices_paid.load_excel_to_df(
        sheet_name='Duplicatas a pagar',
        columns_to_use=[0,1,2,3,4,5,6,7,8,9,10] 
    )
    df_invoices_to_pay = invoices_to_pay.load_excel_to_df(
                sheet_name='em_aberto',
                columns_to_use=[0,1,2,3,4,5,6,7,8,9,10] 
            )
    # 3 - Filter invoices open for payment
    open_invoices_df = invoices_paid.filter_open_invoices(
        df=df_invoices_to_pay,
        type_doc_column='Tipo Doc',
        exclude_value='Contas a Pagar',
        account_column='Conta',
        account_code='2.1.01.01.001'
    )  
    
    # 4 Reconcile (remove already paid invoices from the 'to pay' list)
    invoices_to_process_df = invoices_to_pay.reconcile_invoices(
        invoices_to_pay_df=open_invoices_df,
        paid_invoices_df=df_invoices_paid
    )

    # 5 - Convert values the columns in datetime and balance
    invoices_to_process_df = invoices_to_pay.convert_date_df(invoices_to_process_df, date_columns=['Dt Vencimento', 'Dt Lançamento', 'Dt Documento']) 
    invoices_to_process_df = invoices_to_pay.convert_values_df(invoices_to_process_df, value_columns='Saldo')

    # 6 - Filter by due date
    final_invoices_to_pay_df = invoices_to_pay.filter_by_due_date(
        df=invoices_to_process_df,
        due_data_column='Dt Vencimento',
        target_date=target_date, # Current processing date
    )


    return df_invoices_paid, final_invoices_to_pay_df

def run_merge_and_format_excel(df_invoices_paid: pd.DataFrame, 
                               df_invoices_to_pay: pd.DataFrame) -> None:
    
    """Mesclar e formatar os DataFrames de faturas pagas e
      faturas a pagar, exportando o resultado para um arquivo Excel"""
    
    invoice_to_pay = ExcelProcessor(file_path=PATH_INVOICES_TO_PAY)
    updated_paid_invoices_df = invoice_to_pay.concatenate_dfs(df_invoices_paid, df_invoices_to_pay)

    format_excel = ExcelFormater(file_path=PATH_INVOICES_MERGE, 
                                 df_to_export_excel=updated_paid_invoices_df)

    format_excel.export_df_to_excel(sheet_name='em_aberto')
    format_excel.format_excel_columns(
            sheet_name='em_aberto', 
            date_columns=['G', 'H', 'I'], 
            currency_columns=['J']
        )

def run_main() -> None:
    """Função principal que executa o fluxo de automação, processamento e exportação de faturas."""
    logger.info("Starting the main process for invoice management.")
    
    COLUMNS_DF = ['Fornecedor', 'Tipo Doc', 'No Doc SAP','Saldo', 'Dt Vencimento','Dt Lançamento']
    invoices_to_pay = ExcelProcessor(file_path=PATH_INVOICES_TO_PAY) 
    df_exported_from_sap = run_automation_sap(start_date='01/07/2025', end_date='31/07/2025', current_date='24/07/2025')

    export_report_sap_to_excel(df_sap_to_export=df_exported_from_sap) # type: ignore
    df_paid, df_to_pay = run_processor_excel()
    invoices_to_pay.view_df(df=df_paid, name_columns_load= COLUMNS_DF) # type: ignore
    invoices_to_pay.view_df(df=df_to_pay, name_columns_load= COLUMNS_DF) # type: ignore

if __name__ == '__main__':
    run_main()

