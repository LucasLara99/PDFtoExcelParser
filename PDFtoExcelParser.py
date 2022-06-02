import io
import os
from os import listdir
import pandas as pd
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage


def get_name_files():
   filename = ' '.join(listdir(path='NominaPDF'))
   filenamePath = os.path.join('NominaPDF', filename)
   return filenamePath


def create_df():
   df = pd.DataFrame({'Trabajador': pd.Series(dtype='str'),
                      'Total líquido a percibir': pd.Series(dtype='str')})
   return df


def add_data(df, names, amount):
   for i in range(names.__len__()):
       data = [names[i], amount[i]]
       df.loc[len(df.index)] = data


# FORMAT 1
def get_month_format1(array):
   month = array[array.index('Periodo') + 6]
   return month


def get_names_format1(array):
   names = []
   for i in range(len(array)):
       if array[i] == 'Trabajador:':
           name_end = i
           while array[i] != 'Social:':
               name_start = i
               i -= 1
           name_array = array[name_start:name_end]
           name = ' '.join(name_array)
           names.append(name)
   return names


def get_amounts_format1(array):
   amounts = [array[i + 8] for i in range(len(array)) if array[i] == 'LÍQUIDO']
   return amounts



# FORMAT 2
def get_month_format2(array):
   month = array[array.index('Periodo') - 3]
   return month


def get_names_format2(array):
   names = []
   for i in range(len(array)):
       if array[i] == 'GRAN':
           name_end = i
           while array[i] != 'SL':
               name_start = i
               i -= 1
           name_array = array[name_start:name_end]
           name = ' '.join(name_array)
           names.append(name)
   return names


def get_amounts_format2(array):
   amounts = [array[i - 7] for i in range(len(array)) if array[i] == 'DEDUCIR']
   return amounts


def pdf_to_text(input_file):
   with open(input_file, 'rb') as in_file:
       resMgr = PDFResourceManager()
       retData = io.StringIO()
       TxtConverter = TextConverter(resMgr, retData, laparams=LAParams())
       interpreter = PDFPageInterpreter(resMgr, TxtConverter)
       for page in PDFPage.get_pages(in_file):
           interpreter.process_page(page)

       txt = retData.getvalue()
       array = txt.split()

       if array[4] == 'Empresa:':
           #format1
           month = get_month_format1(array)
           names = get_names_format1(array)
           amounts = get_amounts_format1(array)
       else:
           #format2
           month = get_month_format2(array)
           names = get_names_format2(array)
           amounts = get_amounts_format2(array)


       df = create_df()
       df_with_data = add_data(df, names, amounts)
       writer = pd.ExcelWriter(f'Nominas_{month}.xlsx')
       df.to_excel(writer, sheet_name='Sheet1', index=False, na_rep='NaN')

       for column in df:
           column_width = max(df[column].astype(str).map(len).max(), len(column))
           col_idx = df.columns.get_loc(column)
           writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

       writer.save()


if __name__ == "__main__":
   pdf_to_text(get_name_files())



