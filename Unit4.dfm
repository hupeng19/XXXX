object Form4: TForm4
  Left = 0
  Top = 0
  Caption = 'Form4'
  ClientHeight = 438
  ClientWidth = 822
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  WindowState = wsMaximized
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object dxSpreadSheet1: TdxSpreadSheet
    Left = 113
    Top = 0
    Width = 709
    Height = 438
    Align = alClient
    BorderStyle = cxcbsNone
    Data = {
      C601000044585353763242460700000042465320000000000000000001000101
      01010000000000004246532000000000424653200200000001000000200B0000
      0007000000430061006C00690062007200690000000000002000000020000000
      00200000000020000000002000000000200007000000470045004E0045005200
      41004C0000000000000200000000000000000101000000200B00000007000000
      430061006C006900620072006900000000000020000000200000000020000000
      0020000000002000000000200007000000470045004E004500520041004C0000
      0000000002000000000000000001424653200100000042465320170000005400
      6400780053007000720065006100640053006800650065007400540061006200
      6C00650056006900650077000600000053006800650065007400310001FFFFFF
      FFFFFFFFFF640000000200000002000000020000005500000014000000020000
      0002000000000200000042465320550000000000000042465320000000004246
      5320140000000000000042465320000000000000000000000000010000000000
      0000000000000000000000000000424653200000000000000000000000004246
      53200000000000000000}
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 113
    Height = 438
    Align = alLeft
    Caption = 'Panel1'
    TabOrder = 1
    object Button1: TButton
      Left = 1
      Top = 1
      Width = 111
      Height = 57
      Align = alTop
      Caption = #24320#22987
      TabOrder = 0
      OnClick = Button1Click
    end
    object TreeView1: TTreeView
      Left = 1
      Top = 58
      Width = 111
      Height = 379
      Align = alClient
      AutoExpand = True
      Indent = 19
      TabOrder = 1
      OnDblClick = TreeView1DblClick
    end
  end
  object MainMenu1: TMainMenu
    Left = 160
    Top = 104
    object N1: TMenuItem
      Caption = #25171#24320
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = '  '
      object N3: TMenuItem
        AutoCheck = True
        Caption = '  '#38656#35201#36755#20837#22995#21517
        Checked = True
        OnClick = N3Click
      end
      object N4: TMenuItem
        Caption = #24403#21069#36873#20013#32972#26223#33394
        OnClick = N4Click
      end
      object N5: TMenuItem
        Caption = #24050#32463#36873#36807#30340#32972#26223#33394
        OnClick = N5Click
      end
      object N6: TMenuItem
        AutoCheck = True
        Caption = #33258#21160#25171#24320#32593#39029
        OnClick = N6Click
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = 'excel|*.xlsx'
    Left = 32
    Top = 120
  end
  object ColorDialog1: TColorDialog
    Left = 40
    Top = 192
  end
end
