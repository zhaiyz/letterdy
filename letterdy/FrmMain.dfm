object frmMainFrame: TfrmMainFrame
  Left = 0
  Top = 0
  Caption = #23478#38271#36890#30693#20070#25171#21360
  ClientHeight = 401
  ClientWidth = 650
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object pcMain: TPageControl
    Left = 0
    Top = 0
    Width = 649
    Height = 401
    ActivePage = tsPrint
    TabOrder = 0
    object tsHome: TTabSheet
      Caption = #39318#39029
    end
    object tsTransfer: TTabSheet
      Caption = #25968#25454#36716#31227
      ImageIndex = 1
      object lblExcel: TLabel
        Left = 24
        Top = 24
        Width = 29
        Height = 13
        Caption = 'Excel:'
      end
      object lblSheet: TLabel
        Left = 272
        Top = 24
        Width = 32
        Height = 13
        Caption = 'Sheet:'
      end
      object lblAccess: TLabel
        Left = 16
        Top = 72
        Width = 37
        Height = 13
        Caption = 'Access:'
      end
      object lblTable: TLabel
        Left = 274
        Top = 72
        Width = 30
        Height = 13
        Caption = 'Table:'
      end
      object edtExcel: TEdit
        Left = 59
        Top = 21
        Width = 190
        Height = 21
        ReadOnly = True
        TabOrder = 0
      end
      object edtSheet: TEdit
        Left = 310
        Top = 21
        Width = 121
        Height = 21
        TabOrder = 1
      end
      object btnExcel: TButton
        Left = 448
        Top = 19
        Width = 75
        Height = 25
        Caption = #27983#35272'Excel'
        TabOrder = 2
        OnClick = btnExcelClick
      end
      object edtAccess: TEdit
        Left = 59
        Top = 69
        Width = 190
        Height = 21
        ReadOnly = True
        TabOrder = 3
      end
      object edtTable: TEdit
        Left = 310
        Top = 69
        Width = 121
        Height = 21
        TabOrder = 4
      end
      object btnAccess: TButton
        Left = 448
        Top = 67
        Width = 75
        Height = 25
        Caption = #27983#35272'Access'
        TabOrder = 5
        OnClick = btnAccessClick
      end
      object btnTransfer: TButton
        Left = 544
        Top = 40
        Width = 75
        Height = 33
        Caption = #36827#34892#36716#31227
        TabOrder = 6
        OnClick = btnTransferClick
      end
    end
    object tsPrint: TTabSheet
      Caption = #25968#25454#25171#21360
      ImageIndex = 2
      object lblFirst: TLabel
        Left = 120
        Top = 64
        Width = 52
        Height = 13
        Caption = #31532#19968#25104#32489':'
      end
      object lblSecond: TLabel
        Left = 312
        Top = 64
        Width = 52
        Height = 13
        Caption = #31532#20108#25104#32489':'
      end
      object lblDatabase: TLabel
        Left = 132
        Top = 32
        Width = 40
        Height = 13
        Caption = #25968#25454#24211':'
      end
      object cbbFirst: TComboBox
        Left = 178
        Top = 61
        Width = 119
        Height = 21
        ItemHeight = 13
        TabOrder = 0
      end
      object cbbSecond: TComboBox
        Left = 370
        Top = 61
        Width = 119
        Height = 21
        ItemHeight = 13
        TabOrder = 1
      end
      object edtDatabase: TEdit
        Left = 178
        Top = 29
        Width = 247
        Height = 21
        ReadOnly = True
        TabOrder = 2
      end
      object btnDatabase: TButton
        Left = 431
        Top = 27
        Width = 58
        Height = 25
        Caption = #27983#35272
        TabOrder = 3
        OnClick = btnDatabaseClick
      end
      object btnPrint: TButton
        Left = 431
        Top = 88
        Width = 58
        Height = 25
        Caption = #25171#21360
        TabOrder = 4
        OnClick = btnPrintClick
      end
    end
  end
  object dlgExcel: TOpenDialog
    Filter = 'Excel'#25991#20214'|*.xls'
    Left = 8
    Top = 368
  end
  object dlgAccess: TOpenDialog
    Filter = 'Access'#25991#20214'|*.mdb'
    Left = 40
    Top = 368
  end
  object dlgDatabase: TOpenDialog
    Filter = #25968#25454#24211#65288'Access'#65289'|*.mdb'
    Left = 72
    Top = 368
  end
  object rvdcTable01: TRvDataSetConnection
    FieldAliasList.Strings = (
      #29677'=BAN'
      #32771#21495'=KAOHAO'
      #22995#21517'=XINGMING'
      #24635#20998'=ZONGFEN'
      #26657#21517'=XIAOMING'
      #29677#21517'=BANMING'
      #35821#25991'=YUWEN'
      #25968#23398'=SHUXUE'
      #33521#35821'=YINGYU'
      #29289#29702'=WULI'
      #21270#23398'=HUAXUE'
      #29983#29289'=SHENGWU'
      #25919#27835'=ZHENGZHI'
      #21382#21490'=LISHI'
      #22320#29702'=DILI')
    RuntimeVisibility = rtDeveloper
    DataSet = qryTable01
    Left = 200
    Top = 368
  end
  object rvpMain: TRvProject
    LoadDesigner = True
    ProjectFile = 'E:\delphi\letterdy\report.rav'
    Left = 264
    Top = 368
  end
  object qryTable01: TADOQuery
    Connection = adocMain
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from test02')
    Left = 136
    Top = 368
  end
  object qryTable02: TADOQuery
    Connection = adocMain
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      '')
    Left = 168
    Top = 368
  end
  object adocMain: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\delphi\letterdy\' +
      'data.mdb;Persist Security Info=False'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 104
    Top = 368
  end
  object rvdcTable02: TRvDataSetConnection
    FieldAliasList.Strings = (
      #29677'=BAN'
      #32771#21495'=KAOHAO'
      #22995#21517'=XINGMING'
      #24635#20998'=ZONGFEN'
      #26657#21517'=XIAOMING'
      #29677#21517'=BANMING'
      #35821#25991'=YUWEN'
      #25968#23398'=SHUXUE'
      #33521#35821'=YINGYU'
      #29289#29702'=WULI'
      #21270#23398'=HUAXUE'
      #29983#29289'=SHENGWU'
      #25919#27835'=ZHENGZHI'
      #21382#21490'=LISHI'
      #22320#29702'=DILI')
    RuntimeVisibility = rtDeveloper
    DataSet = qryTable02
    Left = 232
    Top = 368
  end
end