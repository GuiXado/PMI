Attribute VB_Name = "Cadastro"
Option Explicit

Sub Cadastro()
    
    ' Inicia um vetor com 6 índices (0 a 5)
    Dim vetor(5) As Variant
    ' Inicia variáveis paras chamar as planilhas
    Dim planilhaCadastro As Worksheet
    Dim planilhaEstoque As Worksheet
    Dim colunaCadastro As Range
    Dim i%, n%, maior%

    ' Variavel chama a planilha correta a ser manipulada
    Set planilhaCadastro = ThisWorkbook.Sheets("Cadastro")
    Set planilhaEstoque = ThisWorkbook.Sheets("Estoque")
    Set colunaCadastro = planilhaCadastro.Range("A2:G2")
    
    Worksheets("Cadastro").Activate
    
    
    ' Faz as validações
    
    ' Verifica se os campos estão vazios para o cadastro
    For i = 1 To 5
        If planilhaCadastro.Cells(2, i + 1) = "" Then
            MsgBox "Campo em branco!"
            Exit Sub
        End If
    Next i
    
    ' Aqui é pra verificar o tipo de variavel a ser incluida
    Select Case False
    Case TypeName(Range("A2").Value) = "String"
        MsgBox "Campo A2 inválido! O valor deve ser preenchido como texto."
        Exit Sub
    Case IsNumeric(Range("C2").Value)
        MsgBox "Campo C2 inválido! O valor deve ser um número."
        Exit Sub
    Case TypeName(Range("E2").Value) = "String"
        MsgBox "Campo E2 inválido! O valor deve ser preenchido como texto."
        Exit Sub
    End Select

    ' validação do D2 data, recusa produto vencido ou com data igual a atual
    Dim data%
    data = DateDiff("d", Date, planilhaCadastro.Range("D2").Value)
    If data <= 0 Then
        MsgBox "Data de validade inválida. Verifique e tente novamente." ' & vbCrLf & "Campo D2."
        Exit Sub
    End If
    
    ' Se validado..
    
    ' Coloca o conteúdo a ser cadastrado para um vetor
    For i = 0 To 5
        vetor(i) = colunaCadastro.Cells(1, i + 1).Value
    Next i
    
    ' Quando as informações forem para o vetor vou apagar elas da planilha "cadastro"
    planilhaCadastro.Range("A2:F2").ClearContents
    
    ' Coloca uma linha em branco no inicio para que o vetor não cubra nenhuma informação
    planilhaEstoque.Range("A2:I2").Insert
    
    ' Coloca o conteúdo do vetor para a planilha "estoque"
    For i = 0 To 5
        planilhaEstoque.Cells(2, i + 2).Value = vetor(i)
    Next i

    ' gera um código e mostra ele pro usuario, o n° gerado sera o maior número +1
    n = 3
    maior = planilhaEstoque.Range("A3").Value

    Do While planilhaEstoque.Cells(n, "A").Value <> ""
        If planilhaEstoque.Cells(n, "A").Value > maior Then
            maior = planilhaEstoque.Cells(n, "A").Value
        End If
    n = n + 1
    Loop
    planilhaEstoque.Range("A2").Value = maior + 1
    MsgBox "Produtor cadastrado!" & vbCrLf & "Código do produto é: " & maior + 1
    
    ' duplica o estoque cadastrado pro atual, pois esse sera manipilado o outro é só um registro
    planilhaEstoque.Range("I2") = planilhaEstoque.Range("D2")
    
    ' preencher a data do cadastro no estoque
    planilhaEstoque.Range("H2").Value = Date
    
    ' deixa a célula G2 do estoque maiuscula
    planilhaEstoque.Range("F2").Value = UCase(planilhaEstoque.Range("F2").Value)
    planilhaEstoque.Range("G2").Value = UCase(planilhaEstoque.Range("G2").Value)
    planilhaEstoque.Range("B2").Value = UCase(planilhaEstoque.Range("B2").Value)
    
    Dim movi As Variant
    movi = mov

End Sub

Function mov()
    Dim pm As Worksheet, pe As Worksheet
    
    Set pe = ThisWorkbook.Sheets("Estoque")
    Set pm = ThisWorkbook.Sheets("Movimentação")
    
    pm.Range("A2:E2").Insert
    pm.Range("A2:E2").Interior.Color = RGB(235, 241, 222) 'azul claro
    pm.Range("A2") = pe.Cells(2, "A")
    pm.Range("B2") = pe.Cells(2, "B")
    pm.Range("C2") = pe.Cells(2, "E")
    pm.Range("D2") = Date
    pm.Range("E2") = pe.Cells(2, "I")
    'pe.Cells(2, "A").EntireRow.Delete
End Function

