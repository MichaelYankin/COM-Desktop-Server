object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'LR5_Yankin_ExcelClient'
  ClientHeight = 364
  ClientWidth = 660
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 289
    Height = 193
    Caption = #1056#1072#1073#1086#1090#1072' '#1089' '#1076#1086#1082#1091#1084#1077#1085#1090#1086#1084' '
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 34
      Width = 105
      Height = 19
      Caption = #1048#1084#1103' '#1076#1086#1082#1091#1084#1077#1085#1090#1072':'
    end
    object eDocName: TEdit
      Left = 144
      Top = 31
      Width = 129
      Height = 27
      TabOrder = 0
    end
    object bCreateDoc: TButton
      Left = 16
      Top = 72
      Width = 105
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100
      TabOrder = 1
      OnClick = bCreateDocClick
    end
    object bSaveDoc: TButton
      Left = 16
      Top = 103
      Width = 105
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 2
      OnClick = bSaveDocClick
    end
    object bOpenDoc: TButton
      Left = 168
      Top = 72
      Width = 105
      Height = 25
      Caption = #1054#1090#1082#1088#1099#1090#1100
      TabOrder = 3
      OnClick = bOpenDocClick
    end
    object bCloseDoc: TButton
      Left = 168
      Top = 103
      Width = 105
      Height = 25
      Caption = #1047#1072#1082#1088#1099#1090#1100
      TabOrder = 4
      OnClick = bCloseDocClick
    end
    object cbOpen: TCheckBox
      Left = 16
      Top = 134
      Width = 225
      Height = 17
      Caption = #1054#1090#1082#1088#1099#1090#1100' '#1092#1072#1081#1083' '#1087#1088#1080' '#1089#1086#1079#1076#1072#1085#1080#1080
      Checked = True
      State = cbChecked
      TabOrder = 5
    end
    object bConnect: TButton
      Left = 168
      Top = 157
      Width = 105
      Height = 25
      Caption = #1055#1086#1076#1082#1083#1102#1095#1080#1090#1100#1089#1103
      TabOrder = 6
      OnClick = bConnectClick
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 207
    Width = 289
    Height = 146
    Caption = #1052#1077#1090#1086#1076' #1'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    object Label2: TLabel
      Left = 16
      Top = 32
      Width = 265
      Height = 19
      Caption = #1057#1074#1086#1081#1089#1090#1074#1086' Application.FixedDecimalPlaces'
    end
    object Label3: TLabel
      Left = 16
      Top = 57
      Width = 263
      Height = 19
      Caption = #1087#1086#1079#1074#1086#1083#1103#1077#1090' '#1076#1083#1103' '#1083#1080#1089#1090#1072' '#1079#1072#1076#1072#1090#1100' '#1082#1086#1083#1080#1095#1077#1089#1090#1074#1086
    end
    object Label4: TLabel
      Left = 16
      Top = 82
      Width = 259
      Height = 19
      Caption = #1092#1080#1082#1089#1080#1088#1086#1074#1072#1085#1085#1099#1093' '#1076#1077#1089#1103#1090#1080#1095#1085#1099#1093' '#1088#1072#1079#1088#1103#1076#1086#1074':'
    end
    object Label5: TLabel
      Left = 16
      Top = 110
      Width = 109
      Height = 19
      Caption = #1063#1080#1089#1083#1086' '#1088#1072#1079#1088#1103#1076#1086#1074':'
    end
    object ePrecisionValue: TEdit
      Left = 131
      Top = 107
      Width = 47
      Height = 27
      TabOrder = 0
      Text = '2'
    end
    object bSetPrecision: TButton
      Left = 184
      Top = 107
      Width = 89
      Height = 25
      Caption = #1059#1089#1090#1072#1085#1086#1074#1080#1090#1100
      TabOrder = 1
      OnClick = bSetPrecisionClick
    end
  end
  object GroupBox3: TGroupBox
    Left = 303
    Top = 8
    Width = 349
    Height = 193
    Caption = #1052#1077#1090#1086#1076' #2'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    object Label6: TLabel
      Left = 16
      Top = 34
      Width = 290
      Height = 19
      Caption = #1057#1074#1086#1081#1089#1090#1074#1086' Worksheet.StandartWidth '#1086#1090#1074#1077#1095#1072#1077#1090
    end
    object Label7: TLabel
      Left = 16
      Top = 59
      Width = 308
      Height = 19
      Caption = #1079#1072' '#1092#1080#1082#1089#1080#1088#1086#1074#1072#1085#1085#1091#1102' '#1096#1080#1088#1080#1085#1091' '#1089#1090#1086#1083#1073#1094#1072' '#1076#1083#1103' '#1074#1089#1077#1075#1086
    end
    object Label8: TLabel
      Left = 16
      Top = 84
      Width = 267
      Height = 19
      Caption = 'StandartHeight '#1086#1090#1074#1077#1095#1072#1077#1090' '#1079#1072' '#1074#1099#1089#1086#1090#1091' '#1089#1090#1088#1086#1082'.'
    end
    object Label9: TLabel
      Left = 16
      Top = 109
      Width = 314
      Height = 19
      Caption = #1057#1074#1086#1081#1089#1090#1074#1072' '#1087#1086#1079#1074#1086#1083#1103#1102#1090' '#1091#1089#1090#1072#1085#1086#1074#1080#1090#1100' '#1101#1090#1080' '#1079#1085#1072#1095#1077#1085#1080#1103
    end
    object Label13: TLabel
      Left = 16
      Top = 134
      Width = 331
      Height = 19
      Caption = #1076#1083#1103' '#1083#1080#1089#1090#1072', '#1090'.'#1077'. '#1074#1077#1088#1085#1091#1090#1100' '#1080#1093' '#1082' '#1080#1089#1093#1086#1076#1085#1086#1084#1091' '#1079#1085#1072#1095#1077#1085#1080#1102':'
    end
    object bSetHeightAndWidth: TButton
      Left = 232
      Top = 159
      Width = 105
      Height = 25
      Caption = #1059#1089#1090#1072#1085#1086#1074#1080#1090#1100
      TabOrder = 0
      OnClick = bSetHeightAndWidthClick
    end
  end
  object GroupBox4: TGroupBox
    Left = 303
    Top = 207
    Width = 349
    Height = 146
    Caption = #1052#1077#1090#1086#1076' #3'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    object Label10: TLabel
      Left = 16
      Top = 32
      Width = 314
      Height = 19
      Caption = #1052#1077#1090#1086#1076' Chart.SetSourceData '#1087#1086#1079#1074#1086#1083#1103#1077#1090' '#1074#1089#1090#1072#1074#1080#1090#1100
    end
    object Label11: TLabel
      Left = 16
      Top = 57
      Width = 285
      Height = 19
      Caption = #1076#1072#1085#1085#1099#1077', '#1082#1086#1090#1086#1088#1099#1077' '#1085#1072#1093#1086#1076#1103#1090#1089#1103' '#1074' '#1074#1099#1076#1077#1083#1077#1085#1085#1086#1084
    end
    object Label12: TLabel
      Left = 16
      Top = 82
      Width = 165
      Height = 19
      Caption = #1076#1080#1072#1087#1072#1079#1086#1085#1077', '#1074' '#1076#1080#1072#1075#1088#1072#1084#1084#1091':'
    end
    object Label14: TLabel
      Left = 16
      Top = 107
      Width = 111
      Height = 19
      Caption = #1044#1080#1072#1087#1072#1079#1086#1085' A1:C3'
    end
    object bAddChartSource: TButton
      Left = 232
      Top = 107
      Width = 105
      Height = 25
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      TabOrder = 0
      OnClick = bAddChartSourceClick
    end
  end
end
