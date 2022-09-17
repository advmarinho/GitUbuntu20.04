# Layout Contabilização SAP lançamento de rateio

import pandas as pd
import numpy as np

lista = ['31305035','60003', 'CR','Test Contabilização', '19089' ]

campo1 = "E"
campo2 = "0000"
def campo3():
    campo3 = "DB" or "CR"
    if campo3 == lista[2]:
        campo3 = "01"
    else:
        campo3 = "02"
    return campo3
campo3 = campo3()  

print(campo1 + campo2 + campo3)