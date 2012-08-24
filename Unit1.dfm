object Form1: TForm1
  Left = 227
  Top = 133
  Width = 1019
  Height = 616
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object spl1: TSplitter
    Left = 1008
    Top = 0
    Height = 317
    Align = alRight
  end
  object spl2: TSplitter
    Left = 0
    Top = 317
    Width = 1011
    Height = 3
    Cursor = crVSplit
    Align = alBottom
  end
  object pnl1: TPanel
    Left = 0
    Top = 320
    Width = 1011
    Height = 262
    Align = alBottom
    TabOrder = 0
    object spl3: TSplitter
      Left = 505
      Top = 1
      Height = 260
    end
    object strngrd3: TStringGrid
      Left = 1
      Top = 1
      Width = 504
      Height = 260
      Align = alLeft
      FixedCols = 0
      FixedRows = 0
      TabOrder = 0
      ColWidths = (
        64
        194
        52
        51
        48)
    end
    object pnl3: TPanel
      Left = 508
      Top = 1
      Width = 502
      Height = 260
      Align = alClient
      Caption = 'pnl3'
      TabOrder = 1
      object pnlImport: TPanel
        Left = 1
        Top = 1
        Width = 500
        Height = 136
        Align = alTop
        TabOrder = 0
        object lbl1: TLabel
          Left = 8
          Top = 8
          Width = 105
          Height = 13
          Caption = #1058#1072#1073#1083#1080#1094#1072' '#1073#1091#1093#1075#1072#1083#1090#1077#1088#1072':'
        end
        object lbl2: TLabel
          Left = 8
          Top = 32
          Width = 68
          Height = 13
          Caption = #1052#1086#1103' '#1090#1072#1073#1083#1080#1094#1072':'
        end
        object btnBuhTable: TSpeedButton
          Left = 432
          Top = 3
          Width = 23
          Height = 22
          Action = flpnBuhTable
        end
        object btnMyTable: TSpeedButton
          Left = 432
          Top = 28
          Width = 23
          Height = 22
          Action = flpnMyTable
        end
        object edtFromBuh: TEdit
          Left = 123
          Top = 4
          Width = 300
          Height = 21
          TabOrder = 0
        end
        object edtMyTable: TEdit
          Left = 123
          Top = 29
          Width = 300
          Height = 21
          TabOrder = 1
        end
        object chklstCompare: TCheckListBox
          Left = 8
          Top = 56
          Width = 185
          Height = 73
          ItemHeight = 13
          Items.Strings = (
            #1044#1086#1088#1086#1078#1082#1080
            #1052#1077#1090#1072#1083#1083
            #1050#1072#1088#1090#1080#1085#1099
            #1056#1072#1079#1085#1086#1077
            #1050#1086#1074#1088#1099)
          TabOrder = 2
          OnEnter = chklstCompareEnter
        end
      end
    end
  end
  object pnl2: TPanel
    Left = 0
    Top = 0
    Width = 1008
    Height = 317
    Align = alClient
    Caption = 'pnl2'
    TabOrder = 1
    object strngrd2: TStringGrid
      Left = 510
      Top = 1
      Width = 497
      Height = 315
      Align = alRight
      FixedCols = 0
      FixedRows = 0
      TabOrder = 0
      ColWidths = (
        64
        211
        64
        64
        64)
    end
    object strngrd1: TStringGrid
      Left = 1
      Top = 1
      Width = 509
      Height = 315
      Align = alClient
      FixedCols = 0
      RowCount = 11
      FixedRows = 0
      TabOrder = 1
      ColWidths = (
        207
        64
        64
        64
        64)
    end
  end
  object actlst1: TActionList
    Left = 516
    Top = 457
    object flpnBuhTable: TFileOpen
      Category = 'File'
      Caption = '&Open...'
      Dialog.Filter = 'Excel (*.xls)|*.xls'
      Dialog.Title = #1058#1072#1073#1083#1080#1094#1072' '#1073#1091#1093#1075#1072#1083#1090#1077#1088#1072
      Hint = 'Open|Opens an existing file'
      ImageIndex = 7
      ShortCut = 16463
      OnAccept = flpnBuhTableAccept
    end
    object flpnMyTable: TFileOpen
      Category = 'File'
      Caption = '&Open...'
      Dialog.Filter = 'Excel (*.xls)|*.xls'
      Dialog.Title = #1052#1086#1103' '#1090#1072#1073#1083#1080#1094#1072
      Hint = 'Open|Opens an existing file'
      ImageIndex = 7
      ShortCut = 16463
      OnAccept = flpnMyTableAccept
    end
  end
  object il1: TImageList
    Left = 548
    Top = 457
  end
end
