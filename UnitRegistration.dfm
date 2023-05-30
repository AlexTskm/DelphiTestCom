object Form1: TForm1
  Left = 0
  Top = 0
  Caption = #1056#1077#1075#1080#1089#1090#1088#1072#1094#1080#1103' COM '#1086#1073#1098#1077#1082#1090#1086#1074
  ClientHeight = 299
  ClientWidth = 635
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 635
    Height = 41
    Align = alTop
    TabOrder = 0
    object BtnOpen: TButton
      AlignWithMargins = True
      Left = 4
      Top = 4
      Width = 75
      Height = 33
      Hint = #1042#1099#1073#1088#1072#1090#1100' COM '#1086#1073#1098#1077#1082#1090#1099' '#1076#1083#1103' '#1088#1077#1075#1080#1089#1090#1088#1072#1094#1080#1080
      Align = alLeft
      Caption = #1042#1099#1073#1088#1072#1090#1100
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BtnOpenClick
    end
    object BtnReg: TButton
      AlignWithMargins = True
      Left = 85
      Top = 4
      Width = 99
      Height = 33
      Hint = #1056#1077#1075#1080#1089#1090#1088#1080#1088#1086#1074#1072#1090#1100' '#1074#1099#1073#1088#1072#1085#1085#1099#1077' COM '#1086#1073#1098#1077#1082#1090#1099
      Align = alLeft
      Caption = #1056#1077#1075#1080#1089#1090#1088#1080#1088#1086#1074#1072#1090#1100
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BtnRegClick
    end
    object BtnExit: TButton
      AlignWithMargins = True
      Left = 190
      Top = 4
      Width = 75
      Height = 33
      Hint = #1047#1072#1082#1088#1099#1090#1100' '#1087#1088#1080#1083#1086#1078#1077#1085#1080#1077
      Align = alLeft
      Caption = #1042#1099#1093#1086#1076
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BtnExitClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 272
    Width = 635
    Height = 27
    Align = alBottom
    TabOrder = 1
    object Result: TLabel
      Left = 1
      Top = 1
      Width = 633
      Height = 25
      Align = alClient
      ExplicitWidth = 3
      ExplicitHeight = 13
    end
  end
  object ListBox1: TListBox
    Left = 0
    Top = 41
    Width = 635
    Height = 231
    Align = alClient
    ItemHeight = 13
    TabOrder = 2
  end
  object OpenDialog1: TOpenDialog
    Filter = 'COM objects|*.dll'
    Options = [ofHideReadOnly, ofAllowMultiSelect, ofEnableSizing]
    Left = 544
    Top = 72
  end
end
