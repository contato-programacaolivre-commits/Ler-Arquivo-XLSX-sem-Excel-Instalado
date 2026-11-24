object frmPrincipal: TfrmPrincipal
  Left = 0
  Top = 0
  Caption = 'Programa'#231#227'o Livre - Ler Planilha Excel (.XLSX) Puro no Delphi'
  ClientHeight = 470
  ClientWidth = 820
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poScreenCenter
  TextHeight = 15
  object pnlTopo: TPanel
    Left = 0
    Top = 0
    Width = 820
    Height = 97
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    ExplicitWidth = 816
    object lblArquivo: TLabel
      Left = 16
      Top = 12
      Width = 167
      Height = 15
      Caption = '1. Selecione a Planilha (.XLSX):'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblAba: TLabel
      Left = 432
      Top = 12
      Width = 98
      Height = 15
      Caption = '2. Aba da Planilha:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblLinhaInicio: TLabel
      Left = 610
      Top = 12
      Width = 131
      Height = 15
      Caption = '3. Linha In'#237'cio (0=Auto):'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblTotal: TLabel
      Left = 16
      Top = 68
      Width = 137
      Height = 15
      Caption = 'Total de registros lidos: 0'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -12
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object edtArquivo: TEdit
      Left = 16
      Top = 33
      Width = 325
      Height = 23
      ReadOnly = True
      TabOrder = 0
      TextHint = 'Clique em Procurar para escolher um arquivo .xlsx...'
    end
    object btnProcurar: TButton
      Left = 347
      Top = 32
      Width = 75
      Height = 25
      Caption = 'Procurar...'
      TabOrder = 1
      OnClick = btnProcurarClick
    end
    object cbxAbas: TComboBox
      Left = 432
      Top = 33
      Width = 168
      Height = 23
      Style = csDropDownList
      TabOrder = 2
    end
    object speLinhaInicio: TSpinEdit
      Left = 610
      Top = 33
      Width = 85
      Height = 24
      MaxValue = 10000
      MinValue = 0
      TabOrder = 3
      Value = 0
    end
    object btnCarregar: TButton
      Left = 705
      Top = 32
      Width = 99
      Height = 25
      Caption = 'Carregar Grade'
      TabOrder = 4
      OnClick = btnCarregarClick
    end
  end
  object dbgExcel: TDBGrid
    Left = 16
    Top = 103
    Width = 788
    Height = 351
    DataSource = dsExcel
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -12
    TitleFont.Name = 'Segoe UI'
    TitleFont.Style = []
    OnDrawColumnCell = dbgExcelDrawColumnCell
  end
  object dlgAbrirExcel: TOpenDialog
    Filter = 'Planilhas Excel Modernas (*.xlsx)|*.xlsx'
    Title = 'Selecione uma planilha .XLSX'
    Left = 584
    Top = 184
  end
  object memExcel: TFDMemTable
    FetchOptions.AssignedValues = [evMode]
    FetchOptions.Mode = fmAll
    ResourceOptions.AssignedValues = [rvSilentMode]
    ResourceOptions.SilentMode = True
    UpdateOptions.AssignedValues = [uvCheckRequired, uvAutoCommitUpdates]
    UpdateOptions.CheckRequired = False
    UpdateOptions.AutoCommitUpdates = True
    Left = 656
    Top = 184
  end
  object dsExcel: TDataSource
    DataSet = memExcel
    Left = 720
    Top = 184
  end
end
