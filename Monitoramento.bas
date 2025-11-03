Attribute VB_Name = "Monitoramento"
Option Explicit

Sub monitoramento()

    
    Dim matriz() As Variant
    Worksheets("monitoramento").Activate
    Dim planilhaMonitoramento As Worksheet
    Dim planilhaEstoque As Worksheet
    Dim planilhaCadastro As Worksheet
    Dim i%, j%, n&, a#, data%, totalEstoque&, linha&, estoque%, linhaPlan

    Set planilhaMonitoramento = ThisWorkbook.Sheets("Monitoramento")
    Set planilhaEstoque = ThisWorkbook.Sheets("Estoque")
    Set planilhaCadastro = ThisWorkbook.Sheets("Cadastro")

    ' contando quantas linhas possui e talva em total
    totalEstoque = planilhaEstoque.Cells(planilhaEstoque.Rows.Count, "I").End(xlUp).Row
    ' redimensiona a matriz
    ReDim matriz(1 To totalEstoque, 1 To 9)
    
    
    ' aqui ta salvando na matriz, pra depois filtrar
    i = 0 'contador para ver quantos produtos estão a baixo de 50
    
    ' faz as verificações a ser monitoradas e alimenta uma matriz
    For n = 2 To totalEstoque
        data = DateDiff("d", Date, planilhaEstoque.Cells(n, "E").Value)
        'aqui eu verifico se o produto ta com menos de 15% (nao vou deixar um valor fixo)
        If planilhaEstoque.Cells(n, "I").Value <= (planilhaEstoque.Cells(n, "D").Value * 0.2) Or data < 30 Or Date <= 0 Then
            i = i + 1
            For j = 1 To 9
                matriz(i, j) = planilhaEstoque.Cells(n, j).Value
            Next j
        End If
    Next n
    
    ' reseta o padrão da planilha antes de receber as novas informações
    ' e transfere as informações que estão na matriz
    If i > 0 Then
        linhaPlan = planilhaMonitoramento.Cells(planilhaMonitoramento.Rows.Count, 1).End(xlUp).Row + 1
        'limpa a planilha antes do prox "cadastro" digamos
        planilhaMonitoramento.Range("A2:I" & linhaPlan).ClearContents
        'limpa as cores
        planilhaMonitoramento.Range("A2:I" & linhaPlan).Interior.ColorIndex = xlNone
        
        linha = 2
        ' Cola os dados do vetor na planilha de monitoramento
        
        ' aqui ta jogando da matriz pra outra planilha
        For i = 1 To i
            For n = 1 To 9
                planilhaMonitoramento.Cells(linha, n).Value = matriz(i, n)
            Next n
            linha = linha + 1
        Next i
        
    Else
        MsgBox "Não há produtos com estoque abaixo de 50 ou com data inferior a 50 dias."
        
        linhaPlan = planilhaMonitoramento.Cells(planilhaMonitoramento.Rows.Count, 1).End(xlUp).Row + 1
        planilhaMonitoramento.Range("A2:I" & linhaPlan).ClearContents
        planilhaMonitoramento.Range("A2:I" & linhaPlan).Interior.ColorIndex = xlNone
        
    End If
    
    
    'vermelho ambos, estoque o e validade
    'amarelo validade
    'laranja estoque baixo (-15%)
    'preto produto vencido
    
    'sub que pinta as linhas
    
    colorir
    

End Sub

' Faz as verificações para e coloca uma cor em cadas tipo de filtro

Sub colorir()
    Dim pm As Worksheet
    Dim linha#, a#, i%, ultimaColuna%, data%
    
    Set pm = ThisWorkbook.Sheets("Monitoramento")
    linha = pm.Cells(pm.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = 9
    
    For i = 2 To linha
        data = DateDiff("d", Date, pm.Cells(i, "E").Value)
        'Debug.Print data
        If pm.Cells(i, "I").Value <= (pm.Cells(i, "D").Value * 0.2) And data < 30 Then
            pm.Range(pm.Cells(i, 1), pm.Cells(i, ultimaColuna)).Interior.Color = vbRed
            pm.Range(pm.Cells(i, 1), pm.Cells(i, ultimaColuna)).Font.Color = vbBlack
        
        ElseIf data < 0 Then
            pm.Range(pm.Cells(i, 1), pm.Cells(i, ultimaColuna)).Interior.Color = vbBlack
            pm.Range(pm.Cells(i, 1), pm.Cells(i, ultimaColuna)).Font.Color = vbWhite
        
        ElseIf data < 30 Then
            pm.Range(pm.Cells(i, 1), pm.Cells(i, ultimaColuna)).Interior.Color = vbYellow
            pm.Range(pm.Cells(i, 1), pm.Cells(i, ultimaColuna)).Font.Color = vbBlack
            
        ElseIf pm.Cells(i, "I").Value <= (pm.Cells(i, "D").Value * 0.2) Then
            pm.Range(pm.Cells(i, 1), pm.Cells(i, ultimaColuna)).Interior.Color = RGB(255, 165, 0) ' Cor laranja
            pm.Range(pm.Cells(i, 1), pm.Cells(i, ultimaColuna)).Font.Color = vbBlack
        End If
    Next i

End Sub
Sub movimentacao()
    Worksheets("Movimentação").Activate
End Sub

