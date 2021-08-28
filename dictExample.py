from enum import Enum

class colNames(Enum):
    NUMAR_FACTURA = 'Numar factura'
    DATA = 'Data'
    PERSOANA_JURIDICA = 'Persoana juridica'
    CUI = 'CNP / CUI'
    COD_EXTERN = 'Cod extern'
    PRODUS = 'Comenzi'
    COD_PRODUS = 'Cod produs'
    CANTITATE = 'Cantitate'
    PRET = 'Valoare fara TVA'
class colNames2(Enum):
    NUMAR_FACTURA = 'Numar factura'
    DATA = 'Data'
    PERSOANA_JURIDICA = 'Persoana juridica'
    CUI = 'CNP / CUI'
    COD_EXTERN = 'Cod extern'
    PRODUS = 'Comenzi'
    COD_PRODUS = 'Cod produs'
    CANTITATE = 'Cantitate'
    PRET = 'Valoare fara TVA'

s = dict()
s[colNames.NUMAR_FACTURA] = 12
s[colNames.PERSOANA_JURIDICA] = 'DA'
s[colNames2.PERSOANA_JURIDICA] = 'NU'
print(s[colNames.PERSOANA_JURIDICA])