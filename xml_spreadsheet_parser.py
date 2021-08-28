# -*- coding: utf-8 -*-
"""
Created on Wed Aug 18 13:45:57 2021

@author: Sergey.Glukhov
"""

from xml.sax import ContentHandler, parse 
import numpy as np
import pandas as pd

# Reference https://goo.gl/KaOBG3
class ExcelHandler(ContentHandler):
    def __init__(self):
        self.chars = [  ]
        self.cells = [  ]
        self.rows = [  ]
        self.tables = [  ]
    def characters(self, content):
        self.chars.append(content)
    def startElement(self, name, atts):
        if name=="Cell":
            self.chars = [  ]
        elif name=="Row":
            self.cells=[  ]
        elif name=="Table":
            self.rows = [  ]
    def endElement(self, name):
        if name=="Cell":
            self.cells.append(''.join(self.chars))
        elif name=="Row":
            self.rows.append(self.cells)
        elif name=="Table":
            self.tables.append(self.rows)
            

def convert_xml_dataframe(file_name):

# Парсинг данных xml    
    excelHandler = ExcelHandler()
    parse(file_name, excelHandler)
    
# Беру только первую таблицу, предполагая, что все данные на одном листе
    data = excelHandler.tables[0]

# Переупаковываю данные в привычный формат таблицы
    max_len = np.max([len(a) for a in data])
    data_array = np.asarray(
            [np.pad(a, (0, max_len - len(a)), 
                    'constant', constant_values=None) for a in data])
    df = pd.DataFrame(data_array)
    return df

