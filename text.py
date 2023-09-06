import win32com.client
import pandas as pd
import os
import sys
from datetime import timezone,datetime
import ShortName
import ImportDF
import time
import re
from getFileFromWeb import download_csv_files
from getFileFromFolder import copyFiles
import timeit
import util
credentials = util.getCredentials()





pathImportAnalise = '\import\Analise'
columnsAnalise = ['Analise','Status','Site_ID','Rev RF','OBS RF']
Analise = ImportDF.ImportDF_Xlsx(pathImportAnalise,columnsAnalise)
Analise['Analise'] = pd.to_datetime(Analise['Analise'], format="%Y-%m-%d")
Analise['Status'] = Analise['Status'].str.upper()
print(Analise)