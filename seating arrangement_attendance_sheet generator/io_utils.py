import pandas as pd




def read_excel_file(path, logger=None):
    try:
        xls = pd.ExcelFile(path)
    except Exception:
        if logger:
            logger.exception('Unable to open Excel file: %s', path)
        raise


    sheets = {name: xls.parse(name) for name in xls.sheet_names}
    if logger:
        logger.debug('Read sheets: %s', xls.sheet_names)
    return sheets




def write_output_excel(filepath, df):
    df.to_excel(filepath, index=False)