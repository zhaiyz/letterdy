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
    ActivePage = tsHome
    TabOrder = 0
    object tsHome: TTabSheet
      Caption = #39318#39029
    end
    object tsTransfer: TTabSheet
      Caption = #25968#25454#36716#31227
      ImageIndex = 1
      ExplicitLeft = 8
      ExplicitTop = 28
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
  object adocMain: TADOConnection
    Left = 72
    Top = 368
  end
end
