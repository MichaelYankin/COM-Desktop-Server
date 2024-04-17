object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'LR5_Yankin_WordClient'
  ClientHeight = 377
  ClientWidth = 682
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 289
    Height = 209
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
      Top = 32
      Width = 105
      Height = 19
      Caption = #1048#1084#1103' '#1076#1086#1082#1091#1084#1077#1085#1090#1072':'
    end
    object eDocName: TEdit
      Left = 141
      Top = 29
      Width = 132
      Height = 27
      TabOrder = 0
    end
    object bOpenDoc: TButton
      Left = 168
      Top = 72
      Width = 105
      Height = 25
      Caption = #1054#1090#1082#1088#1099#1090#1100
      TabOrder = 1
      OnClick = bOpenDocClick
    end
    object bCreateDoc: TButton
      Left = 16
      Top = 72
      Width = 105
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100
      TabOrder = 2
      OnClick = bCreateDocClick
    end
    object bSaveDoc: TButton
      Left = 16
      Top = 103
      Width = 105
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 3
      OnClick = bSaveDocClick
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
    object bConnect: TButton
      Left = 19
      Top = 166
      Width = 108
      Height = 25
      Caption = #1055#1086#1076#1082#1083#1102#1095#1080#1090#1100#1089#1103
      TabOrder = 5
      OnClick = bConnectClick
    end
    object cbOpen: TCheckBox
      Left = 16
      Top = 143
      Width = 222
      Height = 17
      Caption = #1054#1090#1082#1088#1099#1090#1100' '#1092#1072#1081#1083' '#1087#1088#1080' '#1089#1086#1079#1076#1072#1085#1080#1080
      Checked = True
      State = cbChecked
      TabOrder = 6
    end
    object bFindServer: TButton
      Left = 168
      Top = 166
      Width = 105
      Height = 25
      Caption = #1053#1072#1081#1090#1080' '#1089#1077#1088#1074#1077#1088
      TabOrder = 7
      OnClick = bFindServerClick
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 223
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
      Width = 271
      Height = 19
      Caption = #1052#1077#1090#1086#1076' WordBasic.TableToText '#1087#1086#1079#1074#1086#1083#1103#1077#1090' '
    end
    object Label3: TLabel
      Left = 16
      Top = 57
      Width = 258
      Height = 19
      Caption = #1087#1088#1077#1086#1073#1088#1072#1079#1086#1074#1072#1090#1100' '#1090#1072#1073#1083#1080#1094#1091' '#1074' '#1085#1072#1073#1086#1088' '#1089#1090#1088#1086#1082'.'
    end
    object bConvertTableToText: TButton
      Left = 16
      Top = 107
      Width = 129
      Height = 25
      Caption = #1055#1088#1077#1086#1073#1088#1072#1079#1086#1074#1072#1090#1100
      TabOrder = 0
      OnClick = bConvertTableToTextClick
    end
    object bCreateGenericTable: TButton
      Left = 151
      Top = 107
      Width = 129
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100' '#1090#1072#1073#1083#1080#1094#1091
      TabOrder = 1
      OnClick = bCreateGenericTableClick
    end
  end
  object GroupBox3: TGroupBox
    Left = 303
    Top = 8
    Width = 369
    Height = 209
    Caption = #1052#1077#1090#1086#1076' #2'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    object Label4: TLabel
      Left = 16
      Top = 24
      Width = 303
      Height = 19
      Caption = #1052#1077#1090#1086#1076' WordBasic.DeleteWord '#1091#1076#1072#1083#1103#1077#1090' '#1089#1083#1086#1074#1086' '#1074
    end
    object Label5: TLabel
      Left = 16
      Top = 73
      Width = 202
      Height = 19
      Caption = #1091#1076#1072#1083#1080#1090#1100' '#1080' '#1085#1072#1078#1084#1080#1090#1077' "'#1059#1076#1072#1083#1080#1090#1100'".'
    end
    object Label7: TLabel
      Left = 16
      Top = 49
      Width = 342
      Height = 19
      Caption = #1074#1099#1076#1077#1083#1077#1085#1085#1086#1084' '#1084#1077#1089#1090#1077'. '#1042#1074#1077#1076#1080#1090#1077' '#1089#1083#1086#1074#1086', '#1082#1086#1090#1086#1088#1086#1077' '#1085#1091#1078#1085#1086
    end
    object bDeleteWord: TButton
      Left = 270
      Top = 115
      Width = 88
      Height = 25
      Caption = #1059#1076#1072#1083#1080#1090#1100
      TabOrder = 0
      OnClick = bDeleteWordClick
    end
    object eLineToFind: TEdit
      Left = 16
      Top = 115
      Width = 206
      Height = 27
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Times New Roman'
      Font.Style = [fsItalic]
      ParentFont = False
      TabOrder = 1
      Text = #1042#1074#1077#1076#1080#1090#1077' '#1089#1090#1088#1086#1082#1091'...'
    end
    object cbDeleteAll: TCheckBox
      Left = 16
      Top = 148
      Width = 241
      Height = 17
      Caption = #1059#1076#1072#1083#1080#1090#1100' '#1074#1089#1077' '#1074#1089#1090#1088#1077#1095#1077#1085#1085#1099#1077' '#1089#1083#1086#1074#1072
      TabOrder = 2
    end
  end
  object GroupBox4: TGroupBox
    Left = 303
    Top = 223
    Width = 369
    Height = 146
    Caption = #1052#1077#1090#1086#1076' #3'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    object Label8: TLabel
      Left = 16
      Top = 32
      Width = 272
      Height = 19
      Caption = #1052#1077#1090#1086#1076' WordBasic.InsertPicture '#1087#1086#1079#1074#1086#1083#1103#1077#1090' '
    end
    object Label9: TLabel
      Left = 16
      Top = 57
      Width = 347
      Height = 19
      Caption = #1074#1089#1090#1072#1074#1080#1090#1100' '#1074' '#1076#1086#1082#1091#1084#1077#1085#1090' '#1082#1072#1088#1090#1080#1085#1082#1091'. '#1050#1072#1088#1090#1080#1085#1082#1072' '#1085#1072#1093#1086#1076#1080#1090#1089#1103
    end
    object Label6: TLabel
      Left = 16
      Top = 82
      Width = 238
      Height = 19
      Caption = #1074' '#1087#1072#1087#1082#1077' '#1089' '#1087#1088#1086#1077#1082#1090#1086#1084' '#1076#1072#1085#1085#1086#1081' '#1092#1086#1088#1084#1099'.'
    end
    object bInsertPictureEOF: TButton
      Left = 224
      Top = 107
      Width = 131
      Height = 25
      Caption = #1042#1089#1090#1072#1074#1080#1090#1100' '#1074' '#1082#1086#1085#1077#1094
      TabOrder = 0
      OnClick = bInsertPictureEOFClick
    end
    object bInsertPicture: TButton
      Left = 16
      Top = 107
      Width = 169
      Height = 25
      Caption = #1042#1089#1090#1072#1074#1080#1090#1100' '#1087#1086#1089#1083#1077' '#1082#1091#1088#1089#1086#1088#1072
      TabOrder = 1
      OnClick = bInsertPictureEOFClick
    end
  end
end
