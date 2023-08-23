Sub CopyDataWithDynamicSpacing()
     Application.ScreenUpdating = False ' Desativa a atualização da interface Gráfica
        
    Dim wsOriginal As Worksheet
    Dim wsResultado As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim UltimaLinhaDados As Long
    Dim newRow As Long, currentRow As Long
    Dim destRow As Long
    Dim i As Long, j As Long, k As Long
    Dim novaAba As Worksheet
    Dim ws As Worksheet
    Dim wsHorasProject As Worksheet
    Dim dados As Variant
    Dim resultado As Variant
    Dim busca As String
    Set wsHorasProject = ThisWorkbook.Sheets("Horas_Project")
    
' Copiar a planilha "Horas_Project" e renomeá-la para "HP_BI"
    wsHorasProject.Copy After:=wsHorasProject
    Set wsHP_BI = wsHorasProject.Next
    wsHP_BI.Name = "HP_BI"

' Adicionar uma nova planilha e renomeá-la para "Empilhadas"
    Dim wsEmpilhadas As Worksheet
    Set wsEmpilhadas = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsEmpilhadas.Name = "Empilhadas"
 
' Defina as planilhas de origem e destino
    Set wsOriginal = ThisWorkbook.Sheets("HP_BI")
    Set wsResultado = ThisWorkbook.Sheets("Empilhadas")
    
' Encontre a ultima linha e coluna da planilha original
    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row - 1
    lastCol = wsOriginal.Cells(1, wsOriginal.Columns.Count).End(xlToLeft).Column
    
' Limpar formatação das celulas
    Sheets("HP_BI").Select
    wsOriginal.Cells.Select
    Selection.ClearFormats
    
        'Arruma celulas da primeira coluna
        With wsOriginal
            Dim UltimaLinha As Long
            UltimaLinha = .Cells(.Rows.Count, "A").End(xlUp).Row
      
            For k = 2 To UltimaLinha ' Começa da segunda linha (assumindo que a primeira linha tem cabeçalho)
                .Cells(k, "A").Value = Trim(.Cells(k, "A").Value)
            Next k
        End With
           
    'Criar Colunas
    
    wsOriginal.Select
    Columns("D:D").Select
    Selection.Insert Shift:=x1ToRight
    
'Renomeia as celulas
    wsOriginal.Cells(2, 3).Value = "Sem Alocação"
    wsOriginal.Cells(1, 4).Value = "Nomes"
    
'Copiar nome AA para DD
    i = 0
    For i = i + 1 To lastRow
        If Cells(i, 2) = "" Then
        wsOriginal.Cells(i, 4).Value = wsOriginal.Cells(i, 1).Value
        End If
    Next i
    
    Set ws = ThisWorkbook.Sheets("HP_BI") ' Defina a planilha de trabalho
    currentRow = 2 ' Inicialize a variável de destino da linha
    
'Loop pelas linhas da coluna C
        For destRow = 2 To lastRow + 1
            If Not IsEmpty(ws.Cells(currentRow, "C")) Then ' Verifique se há um dado na célula atual
                    Do While Not IsEmpty(ws.Cells(destRow, "C")) ' Encontre a próxima célula vazia nas células de destino
                        destRow = destRow + 1
                    Loop
                ws.Cells(destRow - 1, "C").Copy ws.Cells(destRow, "C") ' Copie o dado para a próxima célula vazia
            End If
        Next destRow
    
' Inicialize a variável de destino da linha
    destRow = 2
    currentRow = 2
    
' Loop pelas linhas da coluna D
        For destRow = 2 To lastRow + 1
            If Not IsEmpty(ws.Cells(currentRow, "D")) Then ' Verifique se há um dado na célula atual
                Do While Not IsEmpty(ws.Cells(destRow, "D")) ' Encontre a próxima célula vazia nas células de destino
                    destRow = destRow + 1
                Loop
                ws.Cells(destRow - 1, "D").Copy ws.Cells(destRow, "D") ' Copie o dado para a próxima célula vazia
            End If
        Next destRow
    
    'Limpa
    wsOriginal.Select
    Range("A1").Select
    Selection.AutoFilter
    wsOriginal.Range("$A$1:$J$" & i).AutoFilter Field:=2, Criteria1:="="
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
    Application.CutCopyMode = False
        
'Começa a criar as colunas mês
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column ' Última coluna não vazia
    colCount = (lastCol + 1 - 5) 'Calcula a quantidade de colunas de meses já existentes
    col = 5 'Inicia na col 5
    
    For i = colCount + 1 To colCount + colCount ' Adicionando colunas para cada mês adicional
        Columns(col).Insert Shift:=xlToRight
        Cells(1, col).Value = "Mês"
        Cells(2, col).Value = i - colCount
        Range(Cells(2, col), Cells(906, col)).FillDown
        col = col + 2 ' Avançando duas colunas para a próxima iteração
    Next i
    
    Columns("E:AE").EntireColumn.AutoFit 'Ajusta Tamanho Colunas
    
' Encontre a ultima linha e coluna da planilha original
    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row - 1
    lastCol = wsOriginal.Cells(1, wsOriginal.Columns.Count).End(xlToLeft).Column
    UltimaLinhaDados = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row
    
    i = -1
    j = 4
    For i = i To (lastCol - 4) / 2
    For j = j To lastCol - 2
    i = i + 1
    j = j + 1
    
    'colagem do cabecalho
        
'Copiar dados da planilha original
    wsOriginal.Range(wsOriginal.Cells(2, 1), wsOriginal.Cells(lastRow + 1, 4)).Copy
    'Colar dados na planilha de resultados
    wsResultado.Cells((lastRow * i) + 1, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False ' Limpar a área de transferência

'Preencha os dados na planilha de resultados
    wsResultado.Range(wsResultado.Cells((lastRow * i) + 1, 5), wsResultado.Cells((lastRow * (i + 1)), 6)).Value = wsOriginal.Range(wsOriginal.Cells(2, j), wsOriginal.Cells(UltimaLinhaDados, j + 1)).Value

    Next j
    Next i
    
'Cria cabeçalho linha 1

    wsResultado.Rows("1:1").Insert Shift:=xlDown 'Insere 1 linha para o cabeçalho
    wsOriginal.Range("A1:E1").Copy Destination:=wsResultado.Range("A1")
    wsResultado.Cells(1, 6).Value = "Valores"
    wsResultado.Cells(1, 7).Value = "Para"
    Columns("A:AE").EntireColumn.AutoFit 'Ajusta Tamanho Colunas
   
'Realiza Procv dos dados DE/PARA
   lastRow = wsResultado.Cells(wsResultado.Rows.Count, "B").End(xlUp).Row

'Cria tabela para os dados do Proc
    dados = Array( _
        Array("Doppler TruGRD", "PED6263"), _
        Array("Framework 2023", "PED6267"), _
        Array("Monitoramento PRF", "PED9502"), _
        Array("Semex", "PED6280"), _
        Array("Responsividade", "PED6265"), _
        Array("2023 - Redução de Custo", "PED6255"), _
        Array("INVID Fase 2", "PED6258"), _
        Array("Solution Center", "SOC1000"), _
        Array("Plataforma Inova", "TAP300"), _
        Array("Customização Edital CET-SP", "PED6266"), _
        Array("Internacional 2023", "PED6280"), _
        Array("LISA", "TEI1002"), _
        Array("Radar Low Cost Doppler", "PED6281"), _
        Array("Radar Low Cost Laço", "PED6282") _
    )

    For i = 2 To lastRow
        
        busca = wsResultado.Cells(i, 2).Value ' Valor valores na coluna B
        resultado = Application.VLookup(busca, dados, 2, False)
            
            If Not IsError(resultado) Then
                wsResultado.Cells(i, 7).Value = resultado
            End If
    Next i


      
Application.ScreenUpdating = True ' Reativa a Atualização da interface gráfica
    
End Sub


