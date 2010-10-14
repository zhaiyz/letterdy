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
      ExplicitLeft = -60
      ExplicitTop = -64
      ExplicitWidth = 281
      ExplicitHeight = 165
    end
    object tsTransfer: TTabSheet
      Caption = #25968#25454#36716#31227
      ImageIndex = 1
      ExplicitWidth = 281
      ExplicitHeight = 165
    end
    object tsPrint: TTabSheet
      Caption = #25968#25454#25171#21360
      ImageIndex = 2
    end
  end
end
