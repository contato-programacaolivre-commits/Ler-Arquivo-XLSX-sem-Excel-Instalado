unit uLeitorXLSX;

interface

uses
  System.SysUtils, System.Classes, System.Zip, Data.DB,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, Vcl.DBGrids;

type
  TStringArray = array of string;

  TLeitorXLSX = class
  private
    class function LetraParaColuna(const ARef: string): Integer;
    class function DecodificarXML(const ATexto: string): string;
    class function FormatarNumeroExcel(const AValorXML: string): string;
    class function LerAtributo(const ATag, ANome: string): string;
    class procedure CarregarDicionarioTextos(const AXML: string;
      ASaida: TStrings);
  public
    class procedure ListarAbas(const AArquivo: string; ASaida: TStrings);
    class procedure CarregarParaMemTable(const AArquivo: string;
      AIndiceAba: Integer; ADataSet: TFDMemTable; ALinhaInicio: Integer = 1;
      APrimeiraLinhaCabecalho: Boolean = True);
    class procedure AjustarLarguraColunasGrid(AGrid: TDBGrid);
  end;

implementation

class function TLeitorXLSX.LetraParaColuna(const ARef: string): Integer;
var
  I: Integer;
  lLetra: Char;
begin
  Result := 0;
  for I := 1 to Length(ARef) do
  begin
    lLetra := UpCase(ARef[I]);
    if (lLetra >= 'A') and (lLetra <= 'Z') then
      Result := (Result * 26) + (Ord(lLetra) - Ord('A') + 1)
    else
      Break;
  end;
  Dec(Result);
end;

class function TLeitorXLSX.DecodificarXML(const ATexto: string): string;
begin
  Result := ATexto;
  if Pos('&', Result) > 0 then
  begin
    Result := StringReplace(Result, '&amp;', '&', [rfReplaceAll]);
    Result := StringReplace(Result, '&lt;', '<', [rfReplaceAll]);
    Result := StringReplace(Result, '&gt;', '>', [rfReplaceAll]);
    Result := StringReplace(Result, '&quot;', '"', [rfReplaceAll]);
    Result := StringReplace(Result, '&apos;', '''', [rfReplaceAll]);
  end;
end;

class function TLeitorXLSX.FormatarNumeroExcel(const AValorXML: string): string;
var
  lFormatSettings: TFormatSettings;
  lNumero: Extended;
begin
  Result := AValorXML;
  if (Pos('E+', UpperCase(AValorXML)) > 0) or
    (Pos('E-', UpperCase(AValorXML)) > 0) then
  begin
    lFormatSettings := TFormatSettings.Create;
    lFormatSettings.DecimalSeparator := '.';
    if TryStrToFloat(AValorXML, lNumero, lFormatSettings) then
    begin
      if Frac(lNumero) = 0 then
        Result := FormatFloat('0', lNumero)
      else
        Result := FloatToStr(lNumero);
    end;
  end;
end;

class function TLeitorXLSX.LerAtributo(const ATag, ANome: string): string;
var
  lBusca: string;
  lIni, lFim: Integer;
begin
  Result := '';
  lBusca := ANome + '="';
  lIni := Pos(lBusca, ATag);
  if lIni > 0 then
  begin
    Inc(lIni, Length(lBusca));
    lFim := Pos('"', ATag, lIni);
    if lFim > lIni then
      Result := Copy(ATag, lIni, lFim - lIni);
  end;
end;

class procedure TLeitorXLSX.ListarAbas(const AArquivo: string;
  ASaida: TStrings);
var
  lZip: TZipFile;
  lBytes: TBytes;
  lXML, lTag, lNome: string;
  lIdx, lPos, lFim: Integer;
begin
  if not FileExists(AArquivo) then
    raise Exception.Create('Arquivo .xlsx não encontrado: ' + AArquivo);

  ASaida.Clear;
  lZip := TZipFile.Create;
  try
    lZip.Open(AArquivo, zmRead);
    lIdx := lZip.IndexOf('xl/workbook.xml');
    if lIdx < 0 then
      raise Exception.Create
        ('Estrutura interna inválida: xl/workbook.xml não encontrado.');

    lZip.Read(lIdx, lBytes);
    lXML := TEncoding.UTF8.GetString(lBytes);
    // Normaliza prefixo x: caso a planilha tenha sido gerada por bibliotecas externas
    lXML := StringReplace(lXML, '<x:', '<', [rfReplaceAll]);
    lXML := StringReplace(lXML, '</x:', '</', [rfReplaceAll]);

    lPos := 1;
    while lPos > 0 do
    begin
      lPos := Pos('<sheet ', lXML, lPos);
      if lPos > 0 then
      begin
        lFim := Pos('>', lXML, lPos);
        if lFim > lPos then
        begin
          lTag := Copy(lXML, lPos, lFim - lPos + 1);
          lNome := DecodificarXML(LerAtributo(lTag, 'name'));
          if lNome <> '' then
            ASaida.Add(lNome);
          lPos := lFim + 1;
        end
        else
          Break;
      end;
    end;
  finally
    lZip.Free;
  end;
end;

class procedure TLeitorXLSX.CarregarDicionarioTextos(const AXML: string;
  ASaida: TStrings);
var
  lXML, lTextoItem, lTrecho: string;
  lPosSiIni, lPosSiFim, lPosTIni, lPosTFim: Integer;
begin
  ASaida.Clear;
  if AXML = '' then
    Exit;

  lXML := StringReplace(AXML, '<x:', '<', [rfReplaceAll]);
  lXML := StringReplace(lXML, '</x:', '</', [rfReplaceAll]);

  lPosSiIni := 1;
  while lPosSiIni > 0 do
  begin
    lPosSiIni := Pos('<si>', lXML, lPosSiIni);
    if lPosSiIni = 0 then
      lPosSiIni := Pos('<si ', lXML, lPosSiIni);
    if lPosSiIni = 0 then
      Break;

    lPosSiFim := Pos('</si>', lXML, lPosSiIni);
    if lPosSiFim = 0 then
      Break;

    lTextoItem := '';
    lPosTIni := Pos('<t', lXML, lPosSiIni);
    while (lPosTIni > 0) and (lPosTIni < lPosSiFim) do
    begin
      lPosTIni := Pos('>', lXML, lPosTIni);
      if (lPosTIni > 0) and (lPosTIni < lPosSiFim) then
      begin
        Inc(lPosTIni);
        lPosTFim := Pos('</t>', lXML, lPosTIni);
        if (lPosTFim > 0) and (lPosTFim < lPosSiFim) then
        begin
          lTrecho := Copy(lXML, lPosTIni, lPosTFim - lPosTIni);
          lTextoItem := lTextoItem + DecodificarXML(lTrecho);
          lPosTIni := Pos('<t', lXML, lPosTFim + 4);
        end
        else
          Break;
      end
      else
        Break;
    end;

    ASaida.Add(lTextoItem);
    lPosSiIni := lPosSiFim + 5;
  end;
end;

class procedure TLeitorXLSX.CarregarParaMemTable(const AArquivo: string;
  AIndiceAba: Integer; ADataSet: TFDMemTable; ALinhaInicio: Integer;
  APrimeiraLinhaCabecalho: Boolean);
const
  MAX_COLUNAS_SEGURAS = 150;
var
  lZip: TZipFile;
  lBytesShared, lBytesSheet: TBytes;
  lXMLSheet, lLinhaXML, lTagRow, lTagC, lTipo, lValorBruto, lValorFinal,
    lNomeCampo, lCaminhoAba: string;
  lPosCursor, lPosRowIni, lPosRowFim, lFimTagRow, lPosC, lFimC, lPosVIni,
    lPosVFim: Integer;
  lCol, lMaxCol, lIdxTexto, lLinhaAlvo, lNumLinhaAtual, K: Integer;
  lTextos: TStringList;
  lValores: TStringArray;
  lCabecalhoCriado, lLinhaTemDados: Boolean;
begin

  if not FileExists(AArquivo) then
    raise Exception.Create('Arquivo .xlsx não encontrado: ' + AArquivo);

  if ALinhaInicio < 1 then
    lLinhaAlvo := 1
  else
    lLinhaAlvo := ALinhaInicio;

  ADataSet.Close;
  ADataSet.FieldDefs.Clear;

  lZip := TZipFile.Create;
  lTextos := TStringList.Create;
  try

    lZip.Open(AArquivo, zmRead);

    // 1. Lê o dicionário global de strings (xl/sharedStrings.xml)
    if lZip.IndexOf('xl/sharedStrings.xml') >= 0 then
    begin
      lZip.Read('xl/sharedStrings.xml', lBytesShared);
      CarregarDicionarioTextos(TEncoding.UTF8.GetString(lBytesShared), lTextos);
      SetLength(lBytesShared, 0);
    end;

    // 2. Lê o XML da aba escolhida e normaliza eventuais prefixos <x:
    lCaminhoAba := Format('xl/worksheets/sheet%d.xml', [AIndiceAba + 1]);
    if lZip.IndexOf(lCaminhoAba) < 0 then
      raise Exception.Create ('Aba selecionada não encontrada dentro do arquivo .xlsx.');

    lZip.Read(lCaminhoAba, lBytesSheet);
    lXMLSheet := TEncoding.UTF8.GetString(lBytesSheet);
    SetLength(lBytesSheet, 0);

    lXMLSheet := StringReplace(lXMLSheet, '<x:', '<', [rfReplaceAll]);
    lXMLSheet := StringReplace(lXMLSheet, '</x:', '</', [rfReplaceAll]);

    SetLength(lValores, MAX_COLUNAS_SEGURAS);
    lCabecalhoCriado := False;
    lMaxCol := 0;
    lNumLinhaAtual := 0;
    lPosCursor := 1;

    // 3. Percorre o XML por cursor numérico (sem realocar strings gigantes com Delete)
    while lPosCursor > 0 do
    begin
      lPosRowIni := Pos('<row', lXMLSheet, lPosCursor);

      if lPosRowIni = 0 then
        Break;

      lPosRowFim := Pos('</row>', lXMLSheet, lPosRowIni);
      if lPosRowFim = 0 then
        Break;

      lLinhaXML := Copy(lXMLSheet, lPosRowIni, lPosRowFim - lPosRowIni);
      lPosCursor := lPosRowFim + 6;

      // Identifica o número real da linha no Excel (<row r="42">) e pula linhas anteriores
      Inc(lNumLinhaAtual);
      lFimTagRow := Pos('>', lLinhaXML);
      if lFimTagRow = 0 then
        Continue;

      lTagRow := Copy(lLinhaXML, 1, lFimTagRow);
      lNumLinhaAtual := StrToIntDef(LerAtributo(lTagRow, 'r'), lNumLinhaAtual);

      if lNumLinhaAtual < lLinhaAlvo then
        Continue;

      for K := 0 to MAX_COLUNAS_SEGURAS - 1 do
        lValores[K] := '';
      lLinhaTemDados := False;

      // 4. Varre cada célula <c r="A1" t="s"> dentro da linha atual
      lPosC := lFimTagRow + 1;
      while lPosC > 0 do
      begin
        lPosC := Pos('<c ', lLinhaXML, lPosC);
        if lPosC = 0 then
          Break;

        lFimC := Pos('>', lLinhaXML, lPosC);
        if lFimC = 0 then
          Break;

        lTagC := Copy(lLinhaXML, lPosC, lFimC - lPosC + 1);
        lCol := LetraParaColuna(LerAtributo(lTagC, 'r'));

        // Ignora células auto-fechadas vazias (<c r="B2"/>)
        if (lLinhaXML[lFimC - 1] <> '/') and (lCol >= 0) and (lCol < MAX_COLUNAS_SEGURAS) then
        begin
          lTipo := LerAtributo(lTagC, 't');
          lValorFinal := '';

          if lTipo = 'inlineStr' then
          begin
            lPosVIni := Pos('<t', lLinhaXML, lFimC);
            if lPosVIni > 0 then
            begin
              lPosVIni := Pos('>', lLinhaXML, lPosVIni) + 1;
              lPosVFim := Pos('</t>', lLinhaXML, lPosVIni);
              if lPosVFim > lPosVIni then
                lValorFinal :=   DecodificarXML(Copy(lLinhaXML, lPosVIni, lPosVFim - lPosVIni));
            end;
          end
          else
          begin
            lPosVIni := Pos('<v>', lLinhaXML, lFimC);
            if lPosVIni > 0 then
            begin
              Inc(lPosVIni, 3);
              lPosVFim := Pos('</v>', lLinhaXML, lPosVIni);
              if lPosVFim > lPosVIni then
              begin
                lValorBruto := Copy(lLinhaXML, lPosVIni, lPosVFim - lPosVIni);
                if lTipo = 's' then
                begin
                  lIdxTexto := StrToIntDef(lValorBruto, -1);
                  if (lIdxTexto >= 0) and (lIdxTexto < lTextos.Count) then
                    lValorFinal := lTextos[lIdxTexto];
                end
                else
                  lValorFinal := FormatarNumeroExcel
                    (DecodificarXML(lValorBruto));
              end;
            end;
          end;
          if Trim(lValorFinal) <> '' then
          begin
            lValores[lCol] := lValorFinal;
            lLinhaTemDados := True;
            if (not lCabecalhoCriado) and ((lCol + 1) > lMaxCol) then
              lMaxCol := lCol + 1;
          end;
        end;
        lPosC := lFimC + 1;
      end;

      if not lLinhaTemDados then
        Continue;

      // 5. Cria os campos do TFDMemTable usando a linha inicial da tabela
      if not lCabecalhoCriado then
      begin

        for K := 0 to lMaxCol - 1 do
        begin
          lNomeCampo := Trim(lValores[K]);

          if lNomeCampo = '' then
            lNomeCampo := Format('COL_%d', [K + 1]);

          lNomeCampo := Copy(lNomeCampo, 1, 50);

          if ADataSet.FieldDefs.IndexOf(lNomeCampo) >= 0 then
            lNomeCampo := Format('%s_%d', [Copy(lNomeCampo, 1, 40), K + 1]);

          ADataSet.FieldDefs.Add(lNomeCampo, ftString, 255);
        end;

        ADataSet.CreateDataSet;
        ADataSet.DisableControls;

        lCabecalhoCriado := True;
        if not APrimeiraLinhaCabecalho then
        begin
          ADataSet.Append;
          for K := 0 to lMaxCol - 1 do
            ADataSet.Fields[K].AsString := Copy(lValores[K], 1, 255);
          ADataSet.Post;
        end;

      end
      else
      begin
        // 6. Grava os registros das linhas seguintes no TFDMemTable
        ADataSet.Append;
        for K := 0 to lMaxCol - 1 do
          ADataSet.Fields[K].AsString := Copy(lValores[K], 1, 255);
        ADataSet.Post;
      end;
    end;

    if ADataSet.ControlsDisabled then
      ADataSet.EnableControls;

    if ADataSet.Active and (not ADataSet.IsEmpty) then
      ADataSet.First;

  finally
    lTextos.Free;
    lZip.Free;
  end;
end;

class procedure TLeitorXLSX.AjustarLarguraColunasGrid(AGrid: TDBGrid);
var
  I, lLarguraTexto, lLarguraMax, lAmostra: Integer;
  lDataSet: TDataSet;
begin

  if (not Assigned(AGrid)) or (not Assigned(AGrid.DataSource)) or
     (not Assigned(AGrid.DataSource.DataSet)) then
    Exit;

  lDataSet := AGrid.DataSource.DataSet;

  if (not lDataSet.Active) or (AGrid.Columns.Count = 0) then
    Exit;

  lDataSet.DisableControls;
  try
    for I := 0 to AGrid.Columns.Count - 1 do
    begin

      lLarguraMax := AGrid.Canvas.TextWidth(AGrid.Columns[I].Title.Caption) + 24;
      lDataSet.First;

      lAmostra := 0;
      while (not lDataSet.Eof) and (lAmostra < 30) do
      begin
        lLarguraTexto := AGrid.Canvas.TextWidth(lDataSet.Fields[I].AsString) + 20;

        if lLarguraTexto > lLarguraMax then
          lLarguraMax := lLarguraTexto;

        lDataSet.Next;
        Inc(lAmostra);
      end;
      if lLarguraMax < 75 then
        lLarguraMax := 75
      else if lLarguraMax > 260 then
        lLarguraMax := 260;
      AGrid.Columns[I].Width := lLarguraMax;
    end;
    lDataSet.First;
  finally
    lDataSet.EnableControls;
  end;
end;

end.
