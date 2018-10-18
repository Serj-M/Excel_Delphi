object Form1: TForm1
  Left = 404
  Top = 186
  Width = 601
  Height = 353
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object TabbedNotebook1: TTabbedNotebook
    Left = 0
    Top = 0
    Width = 593
    Height = 319
    Align = alClient
    PageIndex = 2
    TabFont.Charset = DEFAULT_CHARSET
    TabFont.Color = clBtnText
    TabFont.Height = -11
    TabFont.Name = 'MS Sans Serif'
    TabFont.Style = []
    TabOrder = 0
    object TTabPage
      Left = 4
      Top = 24
      Caption = #1047#1072#1075#1088#1091#1079#1082#1072' Excel'
      object Button1: TButton
        Left = 24
        Top = 16
        Width = 75
        Height = 25
        Caption = 'Button1'
        TabOrder = 0
        OnClick = Button1Click
      end
      object ListBox1: TListBox
        Left = 120
        Top = 16
        Width = 121
        Height = 97
        ItemHeight = 13
        TabOrder = 1
      end
    end
    object TTabPage
      Left = 4
      Top = 24
      Caption = #1060#1086#1088#1084#1072#1090' '#1103#1095#1077#1081#1082#1080
      object Button2: TButton
        Left = 16
        Top = 8
        Width = 125
        Height = 25
        Caption = #1054#1090#1082#1088#1099#1074#1072#1077#1084' '#1082#1085#1080#1075#1091
        TabOrder = 0
        OnClick = Button2Click
      end
      object Button3: TButton
        Left = 16
        Top = 40
        Width = 125
        Height = 25
        Caption = #1047#1072#1087#1086#1083#1085#1103#1077#1084' '#1103#1095#1077#1081#1082#1080
        Enabled = False
        TabOrder = 1
        OnClick = Button3Click
      end
      object Button10: TButton
        Left = 472
        Top = 256
        Width = 113
        Height = 25
        Caption = #1047#1072#1082#1088#1099#1074#1072#1077#1084' '#1082#1085#1080#1075#1091
        Enabled = False
        TabOrder = 2
        OnClick = Button10Click
      end
      object Button4: TButton
        Left = 16
        Top = 72
        Width = 125
        Height = 25
        Caption = #1063#1080#1090#1072#1077#1084' '#1103#1095#1077#1081#1082#1080
        Enabled = False
        TabOrder = 3
        OnClick = Button4Click
      end
      object Memo1: TMemo
        Left = 160
        Top = 8
        Width = 305
        Height = 281
        Lines.Strings = (
          'Memo1')
        TabOrder = 4
      end
      object Button5: TButton
        Left = 16
        Top = 104
        Width = 125
        Height = 25
        Caption = #1064#1080#1088#1080#1085#1072'/'#1074#1099#1089#1086#1090#1072
        Enabled = False
        TabOrder = 5
        OnClick = Button5Click
      end
      object Button6: TButton
        Left = 16
        Top = 136
        Width = 125
        Height = 25
        Caption = #1060#1086#1088#1084#1072#1090' '#1076#1072#1085#1085#1099#1093
        Enabled = False
        TabOrder = 6
        OnClick = Button6Click
      end
      object Button7: TButton
        Left = 16
        Top = 168
        Width = 125
        Height = 25
        Caption = #1042#1099#1088#1072#1074#1085#1080#1074#1072#1085#1080#1077
        Enabled = False
        TabOrder = 7
        OnClick = Button7Click
      end
      object Button8: TButton
        Left = 16
        Top = 200
        Width = 125
        Height = 25
        Caption = #1055#1086#1074#1086#1088#1086#1090
        Enabled = False
        TabOrder = 8
        OnClick = Button8Click
      end
      object CheckBox1: TCheckBox
        Left = 16
        Top = 228
        Width = 121
        Height = 17
        Caption = #1055#1077#1088#1077#1085#1086#1089' '#1087#1086' '#1089#1083#1086#1074#1072#1084
        Enabled = False
        TabOrder = 9
        OnClick = CheckBox1Click
      end
      object CheckBox2: TCheckBox
        Left = 16
        Top = 246
        Width = 121
        Height = 17
        Caption = #1054#1073#1098#1077#1076#1080#1085#1077#1085#1080#1077' '#1103#1095#1077#1077#1082
        Enabled = False
        TabOrder = 10
        OnClick = CheckBox2Click
      end
      object CheckBox3: TCheckBox
        Left = 16
        Top = 264
        Width = 121
        Height = 17
        Caption = #1040#1074#1090#1086#1087#1086#1076#1073#1086#1088' '#1096#1080#1088#1080#1085#1099
        Enabled = False
        TabOrder = 11
        OnClick = CheckBox3Click
      end
      object Button9: TButton
        Left = 472
        Top = 8
        Width = 113
        Height = 25
        Caption = #1064#1088#1080#1092#1090
        TabOrder = 12
        OnClick = Button9Click
      end
      object Button11: TButton
        Left = 472
        Top = 40
        Width = 113
        Height = 25
        Caption = #1064#1088#1080#1092#1090' +'
        TabOrder = 13
        OnClick = Button11Click
      end
      object Button12: TButton
        Left = 472
        Top = 72
        Width = 113
        Height = 25
        Caption = #1043#1088#1072#1085#1080#1094#1072
        TabOrder = 14
        OnClick = Button12Click
      end
      object Button13: TButton
        Left = 472
        Top = 104
        Width = 113
        Height = 25
        Caption = #1047#1072#1083#1080#1074#1082#1072' '#1103#1095#1077#1081#1082#1080
        TabOrder = 15
        OnClick = Button13Click
      end
    end
    object TTabPage
      Left = 4
      Top = 24
      Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' '#1083#1080#1089#1090#1072
      object Button14: TButton
        Left = 16
        Top = 8
        Width = 125
        Height = 25
        Caption = #1054#1090#1082#1088#1099#1074#1072#1077#1084' '#1082#1085#1080#1075#1091
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        OnClick = Button14Click
      end
      object Button15: TButton
        Left = 16
        Top = 40
        Width = 125
        Height = 25
        Caption = #1047#1072#1087#1086#1083#1085#1103#1077#1084' '#1103#1095#1077#1081#1082#1080
        Enabled = False
        TabOrder = 1
        OnClick = Button15Click
      end
      object Button16: TButton
        Left = 464
        Top = 256
        Width = 113
        Height = 25
        Caption = #1047#1072#1082#1088#1099#1074#1072#1077#1084' '#1082#1085#1080#1075#1091
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        OnClick = Button16Click
      end
      object Button17: TButton
        Left = 16
        Top = 184
        Width = 98
        Height = 25
        Caption = #1055#1088#1086#1089#1084#1086#1090#1088' '#1087#1077#1095#1072#1090#1080
        TabOrder = 3
        OnClick = Button17Click
      end
      object Button18: TButton
        Left = 16
        Top = 216
        Width = 125
        Height = 25
        Caption = #1044#1080#1072#1083#1086#1075' '#1087#1077#1095#1072#1090#1080
        TabOrder = 4
        OnClick = Button18Click
      end
      object Button19: TButton
        Left = 116
        Top = 184
        Width = 25
        Height = 25
        Caption = 'Ex'
        TabOrder = 5
        OnClick = Button19Click
      end
      object Button20: TButton
        Left = 16
        Top = 248
        Width = 125
        Height = 25
        Caption = #1056#1080#1089#1091#1085#1086#1082' '#1092#1086#1085#1072' '#1083#1080#1089#1090#1072
        TabOrder = 6
        OnClick = Button20Click
      end
      object CheckBox4: TCheckBox
        Left = 16
        Top = 72
        Width = 121
        Height = 17
        Caption = #1051#1080#1085#1080#1080' '#1089#1077#1090#1082#1080' '#1103#1095#1077#1077#1082
        Checked = True
        State = cbChecked
        TabOrder = 7
        OnClick = CheckBox4Click
      end
      object CheckBox5: TCheckBox
        Left = 16
        Top = 96
        Width = 129
        Height = 17
        Caption = #1055#1077#1095#1072#1090#1100' '#1083#1080#1085#1080#1081' '#1089#1077#1090#1082#1080
        TabOrder = 8
        OnClick = CheckBox5Click
      end
      object ComboBox1: TComboBox
        Left = 16
        Top = 120
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 9
        OnChange = ComboBox1Change
        Items.Strings = (
          #1050#1085#1080#1078#1085#1072#1103' '#1086#1088#1080#1077#1085#1090#1072#1094#1080#1103
          #1040#1083#1100#1073#1086#1084#1085#1072#1103' '#1086#1088#1080#1077#1085#1090#1072#1094#1080#1103)
      end
      object ComboBox2: TComboBox
        Left = 16
        Top = 144
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 10
        OnChange = ComboBox2Change
        Items.Strings = (
          #1054#1073#1099#1095#1085#1099#1081' '#1074#1080#1076' '#1076#1086#1082#1091#1084#1077#1085#1090#1072
          #1056#1072#1079#1084#1077#1090#1082#1072' '#1089#1090#1088#1072#1085#1080#1094#1099)
      end
    end
  end
  object FontDialog1: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Left = 524
    Top = 32
  end
  object OpenPictureDialog1: TOpenPictureDialog
    Filter = 
      'All (*.jpg;*.jpeg;*.bmp;*.ico;*.emf;*.wmf)|*.jpg;*.jpeg;*.bmp;*.' +
      'ico;*.emf;*.wmf|JPEG Image File (*.jpg)|*.jpg|JPEG Image File (*' +
      '.jpeg)|*.jpeg|Bitmaps (*.bmp)|*.bmp|Enhanced Metafiles (*.emf)|*' +
      '.emf|Metafiles (*.wmf)|*.wmf'
    Left = 196
    Top = 176
  end
end
