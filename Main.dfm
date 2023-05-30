object FMain: TFMain
  Left = 0
  Top = 0
  Caption = #1058#1077#1089#1090#1086#1074#1086#1077' '#1079#1072#1076#1072#1085#1080#1077
  ClientHeight = 299
  ClientWidth = 635
  Color = clBtnFace
  Constraints.MinHeight = 300
  Constraints.MinWidth = 300
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 635
    Height = 25
    Align = alTop
    TabOrder = 0
    object btnCheck: TButton
      Left = 1
      Top = 1
      Width = 120
      Height = 23
      Hint = #1055#1088#1086#1074#1077#1088#1080#1090#1100' '#1087#1072#1088#1072#1084#1077#1090#1088#1099
      Align = alLeft
      Caption = #1055#1088#1086#1074#1077#1088#1080#1090#1100
      Enabled = False
      TabOrder = 0
      OnClick = btnCheckClick
    end
    object btnExecute: TButton
      Left = 121
      Top = 1
      Width = 120
      Height = 23
      Hint = #1042#1099#1087#1086#1083#1085#1080#1090#1100' '#1086#1087#1077#1088#1072#1094#1080#1102
      Align = alLeft
      Caption = #1042#1099#1087#1086#1083#1085#1080#1090#1100
      Enabled = False
      TabOrder = 1
      OnClick = btnExecuteClick
    end
    object btnClearParameters: TButton
      Left = 241
      Top = 1
      Width = 120
      Height = 23
      Hint = #1054#1095#1080#1089#1090#1080#1090#1100' '#1087#1072#1088#1072#1084#1077#1090#1088#1099' '#1086#1087#1077#1088#1072#1094#1080#1080
      Align = alLeft
      Caption = #1054#1095#1080#1089#1090#1080#1090#1100' '#1087#1072#1088#1072#1084#1077#1090#1088#1099
      Enabled = False
      TabOrder = 2
      OnClick = btnClearParametersClick
    end
    object btnExit: TButton
      Left = 361
      Top = 1
      Width = 120
      Height = 23
      Hint = #1047#1072#1082#1088#1099#1090#1100' '#1087#1088#1086#1075#1088#1072#1084#1084#1091
      Align = alLeft
      Caption = #1042#1099#1081#1090#1080
      TabOrder = 3
      OnClick = btnExitClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 271
    Width = 635
    Height = 28
    Align = alBottom
    TabOrder = 1
    object Label1: TLabel
      Left = 1
      Top = 1
      Width = 60
      Height = 26
      Align = alLeft
      Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090': '
      Layout = tlCenter
      ExplicitHeight = 13
    end
    object outResult: TLabel
      Left = 61
      Top = 1
      Width = 3
      Height = 26
      Align = alLeft
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      Layout = tlCenter
      ExplicitHeight = 13
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 25
    Width = 635
    Height = 246
    Align = alClient
    TabOrder = 2
    object GroupBox1: TGroupBox
      Left = 1
      Top = 1
      Width = 633
      Height = 40
      Align = alTop
      Caption = #1044#1077#1081#1089#1090#1074#1080#1077':'
      TabOrder = 0
      object cbAction: TComboBox
        Left = 2
        Top = 15
        Width = 629
        Height = 21
        Hint = #1042#1099#1073#1077#1088#1080#1090#1077' '#1087#1072#1088#1072#1084#1077#1090#1088#1080#1095#1077#1089#1082#1091#1102' '#1086#1087#1077#1088#1072#1094#1080#1102
        Align = alTop
        Style = csDropDownList
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnChange = cbActionChange
      end
    end
    object GroupBox2: TGroupBox
      Left = 1
      Top = 41
      Width = 633
      Height = 204
      Align = alClient
      Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099':'
      TabOrder = 1
    end
  end
end
