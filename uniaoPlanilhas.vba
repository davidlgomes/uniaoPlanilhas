Sub CombinarPlanilhas()
    Dim ws As Worksheet
    Dim novaWs As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim j As Long

    ' Cria uma nova planilha para armazenar os dados combinados
    Set novaWs = ThisWorkbook.Worksheets.Add
    novaWs.Name = "Dados Combinados"

    ' Adiciona os cabeçalhos na nova planilha
    novaWs.Cells(1, 1).Value = "Nome da Planilha"
    novaWs.Cells(1, 2).Value = "Dados"

    ' Inicializa a variável para a linha da nova planilha
    ultimaLinha = 2

    ' Itera por todas as planilhas no workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Ignora a planilha de dados combinados
        If ws.Name <> novaWs.Name Then
            ' Itera por todas as linhas da planilha atual
            For i = 1 To ws.UsedRange.Rows.Count
                ' Copia o nome da planilha
                novaWs.Cells(ultimaLinha, 1).Value = ws.Name

                ' Copia os dados da linha
                For j = 1 To ws.UsedRange.Columns.Count
                    novaWs.Cells(ultimaLinha, j + 1).Value = ws.Cells(i, j).Value
                Next j

                ' Avança para a próxima linha na nova planilha
                ultimaLinha = ultimaLinha + 1
            Next i
        End If
    Next ws

    ' Ajusta a largura das colunas na nova planilha
    novaWs.Columns("A:Z").AutoFit

    MsgBox "Combinação concluída!"
End Sub
