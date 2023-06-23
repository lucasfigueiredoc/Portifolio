
import cx_Oracle
import os, sys
import traceback

import datetime
from datetime import datetime

oracle_pass = os.environ.get('ORACLEPASS')
oracle_user = os.environ.get('ORACLEUSER')

conn = cx_Oracle.connect(f'{oracle_user}/{oracle_pass}@xxxxx')
cur = conn.cursor()

timeStart = datetime.strftime(datetime.now(),'%d/%m/%Y %H:%M:%S') ## Registro do inicio da execução da script para monitoramento
def executarScript(scriptName):
    
    try:
        exec('import {}'.format(scriptName))

        execValue = 1       
        timeEnd = datetime.now().strftime('%Y-%m-%d %H:%M:%S') #datetime.strftime(datetime.now(),'%d/%m/%Y %H:%M:%S') ## Registro do fim da execução da script para monitoramento
        
        query = "INSERT INTO logs_execucao_spy_script(script_name, time_start, time_end, exec_value) VALUES('{}','{}','{}','{}')".format(scriptName,timeStart,timeEnd,execValue)

        cur.execute(query) # Execução da query no BD
        conn.commit()    ## Commit da querry
        
    except :
        exec('import {}'.format(scriptName))
        traceB = str(traceback.format_exc()) ## Armazena o traceback em uma str para registro no bd
        print(traceB)

        execValue = 0
        timeEnd = datetime.now().strftime('%Y-%m-%d %H:%M:%S') #datetime.strftime(datetime.now(),'%d/%m/%Y %H:%M:%S')
        query = "INSERT INTO logs_execucao_spy_script(script_name, time_start, time_end, exec_value) VALUES('{}','{}','{}','{}')".format(scriptName,timeStart,timeEnd,execValue)
        
        cur.execute(query) # Execução da query no BD
        conn.commit() ## Commit da querry


scriptName = sys.argv[1]
executarScript(scriptName)
conn.close()
