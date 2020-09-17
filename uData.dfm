object fData: TfData
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 233
  Top = 199
  Height = 150
  Width = 215
  object Database: TIBDatabase
    DatabaseName = 'E:\'#1055#1088#1086#1075#1088#1072#1084#1084#1080#1088#1086#1074#1072#1085#1080#1077'\nSchool\1.FDB'
    Params.Strings = (
      'lc_ctype=WIN1251'
      'user_name=sysdba'
      'password=masterkey')
    LoginPrompt = False
    DefaultTransaction = Transaction
    IdleTimer = 0
    SQLDialect = 3
    TraceFlags = [tfQExecute, tfError, tfConnect]
    Left = 16
    Top = 8
  end
  object Transaction: TIBTransaction
    Active = False
    DefaultDatabase = Database
    DefaultAction = TACommitRetaining
    AutoStopAction = saNone
    Left = 64
    Top = 8
  end
  object SQL: TIBSQL
    Database = Database
    ParamCheck = True
    Transaction = Transaction
    Left = 112
    Top = 8
  end
  object ExcelA: TExcelApplication
    AutoConnect = False
    ConnectKind = ckNewInstance
    AutoQuit = False
    Left = 152
    Top = 8
  end
  object SQLMonitor: TIBSQLMonitor
    OnSQL = SQLMonitorSQL
    TraceFlags = [tfQExecute, tfError, tfConnect]
    Left = 24
    Top = 64
  end
  object WordA: TWordApplication
    AutoConnect = False
    ConnectKind = ckNewInstance
    AutoQuit = False
    Left = 88
    Top = 64
  end
end
