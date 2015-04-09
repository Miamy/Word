object FMain: TFMain
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'MS Word table generator'
  ClientHeight = 509
  ClientWidth = 563
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object Label1: TLabel
    Left = 16
    Top = 16
    Width = 75
    Height = 14
    Caption = 'Input files dir:'
  end
  object Label2: TLabel
    Left = 16
    Top = 42
    Width = 63
    Height = 14
    Caption = 'Output file:'
  end
  object lblStatus: TLabel
    Left = 16
    Top = 72
    Width = 171
    Height = 14
    Caption = 'Directory not processed yet'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object PictureImage: TImage
    Left = 8
    Top = 104
    Width = 105
    Height = 105
    AutoSize = True
    Center = True
    Proportional = True
    Stretch = True
    Visible = False
  end
  object Label7: TLabel
    Left = 16
    Top = 442
    Width = 48
    Height = 14
    Caption = 'Scale,%:'
  end
  object lblWord: TLabel
    Left = 16
    Top = 104
    Width = 144
    Height = 14
    Caption = 'Connecting to MS Word...'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsItalic]
    ParentFont = False
    Visible = False
  end
  object lblFilesPerPage: TLabel
    Left = 164
    Top = 442
    Width = 171
    Height = 14
    Caption = 'Directory not processed yet'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object edSelectPictures: TEdit
    Left = 120
    Top = 8
    Width = 388
    Height = 22
    TabOrder = 0
    Text = 'd:\Programs\Word\small\'
  end
  object edSelectDoc: TEdit
    Left = 120
    Top = 39
    Width = 388
    Height = 22
    TabOrder = 1
    Text = 'd:\Programs\Word\test.doc'
  end
  object btnSelectPictures: TButton
    Left = 514
    Top = 8
    Width = 25
    Height = 25
    Caption = '...'
    TabOrder = 2
    OnClick = btnSelectPicturesClick
  end
  object btnSelectDoc: TButton
    Left = 514
    Top = 38
    Width = 25
    Height = 25
    Caption = '...'
    TabOrder = 3
    OnClick = btnSelectDocClick
  end
  object btnGenerate: TButton
    Left = 160
    Top = 476
    Width = 137
    Height = 25
    Caption = 'Generate'
    TabOrder = 4
    OnClick = btnGenerateClick
  end
  object btnExit: TButton
    Left = 464
    Top = 476
    Width = 75
    Height = 25
    Caption = 'Exit'
    TabOrder = 5
    OnClick = btnExitClick
  end
  object btnCount: TButton
    Left = 16
    Top = 476
    Width = 75
    Height = 25
    Caption = 'Count'
    TabOrder = 6
    OnClick = btnCountClick
  end
  object GroupBox1: TGroupBox
    Left = 16
    Top = 376
    Width = 523
    Height = 57
    Caption = 'Margins, mm'
    TabOrder = 7
    object Label3: TLabel
      Left = 12
      Top = 27
      Width = 26
      Height = 14
      Caption = 'Left:'
    end
    object Label4: TLabel
      Left = 132
      Top = 27
      Width = 32
      Height = 14
      Caption = 'Right:'
    end
    object Label5: TLabel
      Left = 264
      Top = 27
      Width = 26
      Height = 14
      Caption = 'Top:'
    end
    object Label6: TLabel
      Left = 389
      Top = 27
      Width = 45
      Height = 14
      Caption = 'Bottom:'
    end
    object seLeft: TSpinEdit
      Left = 48
      Top = 24
      Width = 57
      Height = 23
      Increment = 5
      MaxValue = 0
      MinValue = 0
      TabOrder = 0
      Value = 40
    end
    object seRight: TSpinEdit
      Left = 176
      Top = 24
      Width = 57
      Height = 23
      Increment = 5
      MaxValue = 0
      MinValue = 0
      TabOrder = 1
      Value = 25
    end
    object seTop: TSpinEdit
      Left = 304
      Top = 24
      Width = 57
      Height = 23
      Increment = 5
      MaxValue = 0
      MinValue = 0
      TabOrder = 2
      Value = 20
      OnChange = seScaleChange
    end
    object seBottom: TSpinEdit
      Left = 448
      Top = 24
      Width = 57
      Height = 23
      Increment = 5
      MaxValue = 0
      MinValue = 0
      TabOrder = 3
      Value = 20
      OnChange = seScaleChange
    end
  end
  object seScale: TSpinEdit
    Left = 74
    Top = 439
    Width = 57
    Height = 23
    Increment = 10
    MaxValue = 0
    MinValue = 0
    TabOrder = 8
    Value = 100
    OnChange = seScaleChange
  end
  object SaveDialog: TSaveDialog
    DefaultExt = '*.doc'
    Filter = 'Word files|*.doc'
    Left = 384
    Top = 128
  end
  object WordApplication: TWordApplication
    AutoConnect = False
    ConnectKind = ckNewInstance
    AutoQuit = False
    Left = 192
    Top = 136
  end
  object WordDocument: TWordDocument
    AutoConnect = False
    ConnectKind = ckNewInstance
    Left = 288
    Top = 136
  end
end
