unit uPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Data.DB, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, Vcl.Grids, Vcl.DBGrids,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Samples.Spin, uLeitorXLSX;

type
  TfrmPrincipal = class(TForm)
    pnlTopo: TPanel;
    lblArquivo: TLabel;
    lblAba: TLabel;
    lblLinhaInicio: TLabel;
    edtArquivo: TEdit;
    btnProcurar: TButton;
    cbxAbas: TComboBox;
    speLinhaInicio: TSpinEdit;
    btnCarregar: TButton;
    lblTotal: TLabel;
    dbgExcel: TDBGrid;
    dlgAbrirExcel: TOpenDialog;
    memExcel: TFDMemTable;
    dsExcel: TDataSource;
    procedure btnProcurarClick(Sender: TObject);
    procedure btnCarregarClick(Sender: TObject);
    procedure dbgExcelDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private
    procedure ProcessarAbaSelecionada;
  public
  end;

var
  frmPrincipal: TfrmPrincipal;

implementation

{$R *.dfm}

procedure TfrmPrincipal.ProcessarAbaSelecionada;
begin
  if (edtArquivo.Text = '') or (cbxAbas.ItemIndex < 0) then
    Exit;

  Screen.Cursor := crHourGlass;
  try

    TLeitorXLSX.CarregarParaMemTable(edtArquivo.Text, cbxAbas.ItemIndex,
      memExcel, speLinhaInicio.Value, True);

    TLeitorXLSX.AjustarLarguraColunasGrid(dbgExcel);

    lblTotal.Caption := Format('Total de registros lidos: %d', [memExcel.RecordCount]);
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmPrincipal.btnProcurarClick(Sender: TObject);
begin
  if dlgAbrirExcel.Execute then
  begin
    edtArquivo.Text := dlgAbrirExcel.FileName;
    try
      TLeitorXLSX.ListarAbas(edtArquivo.Text, cbxAbas.Items);
      if cbxAbas.Items.Count > 0 then
      begin
        cbxAbas.ItemIndex := 0;
        ProcessarAbaSelecionada;
      end;
    except
      on E: Exception do
        ShowMessage('Erro ao abrir planilha: ' + E.Message);
    end;
  end;
end;

procedure TfrmPrincipal.btnCarregarClick(Sender: TObject);
begin
  try
    ProcessarAbaSelecionada;
  except
    on E: Exception do
      ShowMessage('Erro ao carregar aba: ' + E.Message);
  end;
end;

procedure TfrmPrincipal.dbgExcelDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if not (gdSelected in State) then
  begin
    if Odd(memExcel.RecNo) then
      dbgExcel.Canvas.Brush.Color := clWhite
    else
      dbgExcel.Canvas.Brush.Color := $00F5F7FA;
    dbgExcel.Canvas.Font.Color := clWindowText;
  end;

  dbgExcel.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

end.

