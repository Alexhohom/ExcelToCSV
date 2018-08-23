object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = #20027#30028#38754
  ClientHeight = 328
  ClientWidth = 643
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 263
    Width = 643
    Height = 65
    Align = alBottom
    TabOrder = 0
    object btnToCSV: TButton
      Left = 410
      Top = 6
      Width = 96
      Height = 46
      Caption = #25171#24320#25991#20214
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #24494#36719#38597#40657
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnClick = btnToCSVClick
    end
  end
  object list1: TListBox
    Left = 0
    Top = 0
    Width = 643
    Height = 263
    Align = alClient
    ItemHeight = 13
    TabOrder = 1
  end
  object OpenDialog1: TOpenDialog
    Left = 239
    Top = 158
  end
  object Timer1: TTimer
    OnTimer = Timer1Timer
    Left = 471
    Top = 181
  end
end
