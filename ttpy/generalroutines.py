
from docx import Document

## Replace in docx
###########################
def Replace(replace_dict,doc):
    for item in replace_dict:
        # Iterate through paragraphs and replace the text
        for paragraph in doc.paragraphs:
            if item in paragraph.text:
                paragraph.text = paragraph.text.replace(item, replace_dict[item])
                
    return doc

## Add table in docx
###########################       
def AddTable(target_string,df,doc):    
    from docx.shared import Inches
    # Find the paragraph that contains the target string
    target_paragraph = None

    for paragraph in doc.paragraphs:
        if target_string in paragraph.text:
            target_paragraph = paragraph
            break

    # add a table to the end and create a reference variable
    # extra row is so we can add the header row
    table = doc.add_table(df.shape[0]+1, df.shape[1],style='StijlHans')

    # Insert the table after the target paragraph
    p = target_paragraph._element
    new_tbl = table._tbl
    p.addnext(new_tbl)

    # Remove the target text if needed
    run = target_paragraph.runs[0]
    run.text = run.text.replace(target_string, '')


    # add the header rows.
    for j in range(df.shape[-1]):
        table.cell(0,j).text = df.columns[j]

    # add the rest of the data frame
    for i in range(df.shape[0]):
        for j in range(df.shape[-1]):
            table.cell(i+1,j).text = str(df.values[i,j])
            
    # Apply column widths
    i=0
    for cell in table.columns[0].cells:
        if i==0:
            cell.width = Inches(4)
        else:
            cell.width = Inches(2)
    return doc

def SaveExcel(df, filename, sheetname):
    import pandas as pd

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name=sheetname, index=False, na_rep='')

    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets[sheetname]

    # Add a format to wrap text.
    wrap_format = workbook.add_format({'text_wrap': True})

    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))*1.02
        column_length = min(column_length, 150)
        col_idx = df.columns.get_loc(column)

        # Apply the format to the column.
        worksheet.set_column(col_idx, col_idx, column_length, wrap_format)

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()
"""
## Save in Excel
def SaveExcel(df,filename,sheetname):
    import pandas as pd
    writer = pd.ExcelWriter(filename) 
    df.to_excel(writer, sheet_name=sheetname, index=False, na_rep='')
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))*1.02
        # add total length of the column
        column_length = max(column_length, 100)
        col_idx = df.columns.get_loc(column)
        writer.sheets[sheetname].set_column(col_idx, col_idx, column_length)

    writer.close()
"""
    

    