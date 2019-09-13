from os import PathLike
from os.path import expandvars
from pathlib import Path
from typing import Union, Sequence, List, Tuple

from pandas import DataFrame, ExcelWriter


def to_autosize_excel(df: DataFrame,
                      outfile: PathLike,
                      consider_headers: bool = True,
                      sheet_name: str='Sheet1',
                      na_rep: str='',
                      float_format: str=None,
                      columns: Union[Sequence[str], List[str]]=None,
                      header: Union[bool, List[str]]=True,
                      index: bool=True,
                      index_label: Union[str, Sequence]=None,
                      startrow: int=0,
                      startcol: int=0,
                      inf_rep: str='inf',
                      verbose: bool=True,
                      freeze_panes: Tuple[int,int]=None,
                      excel_date_format: str = "yyyy-mm-dd",
                      excel_datetime_format: str = "yyyy-mm-dd  hh:mm:ss",
                      mode: str='w')-> Path:
    """
    
    Arguments:
        df {DataFrame} -- The data to be output into an xlsx file
        outfile {PathLike} -- A pathlike object representing the full path and filename of the output xlsx file
    
    Keyword Arguments:
        consider_headers {bool} -- If true, consider the width of the column headers when sizing columns (default: {True})
        sheet_name {str} -- The sheet of the workbook to write the data(default: {'Sheet1'})
        na_rep {str} -- How null values should be represented in the output (default: {''})
        float_format {str} -- Format string for floating point numbers. (default: {None})
        columns {Union[Sequence[str], List[str]]} -- If given, only these columns will be written to the file (default: {None})
        header {Union[bool, List[str]]} -- [description] (default: {True})
        index {bool} -- If true, write the index columns in the output (default: {True})
        index_label {Union[str, Sequence]} -- Alternative column headers for index columns. (default: {None})
        startrow {int} -- The zero-indexed row of the xlsx file to begin writing data (default: {0})
        startcol {int} -- The zero-indexed column of the xlsx file to begin writing data (default: {0})
        inf_rep {str} -- How the value of infinity will be represnted in the output (default: {'inf'})
        verbose {bool} -- Display more information in the error logs. (default: {True})
        freeze_panes {Tuple[int,int]} -- Specifies the one-based bottommost row and rightmost column that is to be frozen. (default: {None})
        excel_date_format {str} -- Format string for dates written into Excel files  (default: {"yyyy-mm-dd"})
        excel_datetime_format {str} -- Format string for datetime objects written into Excel files (default: {"yyyy-mm-dd  hh:mm:ss"})
        mode {str} -- Must equal 'w' (write) or 'a' (append)  (default: {'w'})
    
    Returns:
        Path -- A Path object representing the successfully written xlsx output
    """
    #we don't want to pass df or outfile as kwargs later
    kwargs = {k:v for k,v in zip(list(locals().keys())[3:], list(locals().values())[3:])}

    #construct the ExcelWriter, removing its kwargs as they will no longer be needed
    writer = ExcelWriter(str(Path(expandvars(outfile))),
                         engine="xlsxwriter",
                         date_format=kwargs.pop("excel_date_format"),
                         datetime_format=kwargs.pop("excel_datetime_format"),
                         mode=kwargs.pop("mode"))

    #This just makes things easier later, trust me.  Also df is probably mutable, so not even risking screwing it up!
    if kwargs["columns"]:
        data = df[list(kwargs["columns"])]
    else:
        data = df

    with writer:
        #only kwargs left should be kwargs of df.to_excel
        df.to_excel(writer, **kwargs)
        wb = writer.book
        ws = writer.sheets[kwargs["sheet_name"]]

        if isinstance(columns, bool): #Use the DataFrame's existing labels
            '''if also going to write index, mash it into the dataframe and just get the
            index level names as columns'''
            if index:
                labels = data.reset_index().columns.to_list()
            else:
                labels = data.columns.to_list()
        else: #Use provided alternative labels
            if index and index_label: #Use provided index label(s)
                if isinstance(index_label, str):
                    labels = [index_label] + list(columns)
                else:
                    labels = list(index_label) + list(columns)
            elif index and not index_label: #Use existing index name(s) as label(s) with alternative column labels
                #a labeless index has a Nonetype name, which converts to the string "None".  I prefer the empty string.
                labels = [str(name) if name else "" for name in data.index.names] + list(columns)
            else:
                labels = list(columns)

        if index: #much easier to get widths if you just treat the index like regular columns
            widths = maximum_character_widths(data.reset_index(), consider_headers, labels)
        else:
            widths = maximum_character_widths(data, consider_headers, labels)

        #size columns using calculated best-fit widths
        for column in range(startcol, startcol+len(labels)):
            if index:
                column_name = data.reset_index().columns[column]
            else:
                column_name = data.columns[column]
            column_width = widths[column_name]
            ws.set_column(column, column, excel_column_width(column_width))

        #re-write the columns with a custom format that wraps text if columns headers were not considered in sizing of columns
        if columns and not consider_headers:
            f = wb.add_format({"text_wrap":True, "bold":True, "align":"center", "valign":"vcenter", "border":1})
            ws.write_row(startrow, startcol, labels, f)


    return Path(writer.path)

def maximum_character_widths(df: DataFrame, consider_headers: bool = True, alternate_headers: Union[list,dict] = None) -> dict:
    """Gets the maximum character width (i.e. the length of the string) of a column in a dataframe.  Optionally considers the headers when determining the maximum width
    
    Arguments:
        df {DataFrame} -- The input data
    
    Keyword Arguments:
        consider_headers {bool} --  If true, consider the column header when determining maximum width. (default: {True})
        alternate_headers {Union[list,dict]} -- If present, is equivalent to consider_headers = True, except these values will be considered instead of column labels. (default: {None})
    
    Raises:
        ValueError: Raised if the number of alternative column headers does not match the number of columns in the dataframe
        TypeError: Raised if alternative headers is not a list or dictionary
    
    Returns:
        dict -- A dictionary of character widths by column header
    """
    widths = {}

    if isinstance(alternate_headers, list):
        if len(alternate_headers) != len(df.columns):
            raise ValueError("The number of labels must equal the number of columns in the dataframe")
        else:
            headers = {k:v for k,v in zip(df.columns, alternate_headers)}
    elif isinstance(alternate_headers, dict):
        if len(alternate_headers.keys()) != len(df.columns):
            raise ValueError("The number of labels must equal the number of columns in the dataframe")
        else:
            headers = alternate_headers
    elif alternate_headers == None and consider_headers:
        headers = {v:v for v in df.columns}
    else:
        raise TypeError("Alternative headers must be a list or dictionary")

    for key,value in headers.items():
        if consider_headers:
            widths[key] = max(len(value), df[key].astype(str).str.len().max())
        else:
            widths[key] = df[key].astype(str).str.len().max()

    return widths

def excel_column_width(charwidth:int, fontsize:float=11) -> float:
    """Converts a character width to a an Excel column width based on the font size
    
    Arguments:
        charwidth {int} -- The number of characters in the cell value to fit the column to
    
    Keyword Arguments:
        fontsize {float} --  The font size of the cell to fit. (default: {11})
    
    Returns:
        float -- The value of a close-enough Excel column width
    """
    #emperically derived from observation of excel.  At best this is an approximation that errs on the side of slightly oversized
    return charwidth * round(0.118775 * fontsize, 2) 