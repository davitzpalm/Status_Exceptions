from sqlalchemy import create_engine
import urllib


def create_mav_engine():
    server = 'mdr-db-server.database.windows.net,1433'
    db = 'MDR'
    user = 'mdr-readonly-login'
    pwd = '0VSurdvt1Eo7oCoabZzV'

    params = urllib.parse.quote_plus('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+ server +';DATABASE='+ db +';UID='+ user +';PWD='+ pwd)

    engine = create_engine("mssql+pyodbc:///?odbc_connect=%s" % params)
    return engine