Attribute VB_Name = "Consulta"
Option Explicit

Sub consulta() ' é um "menu" das opções de retirada e algumas validações
    Dim n As Variant
    
    Do
        n = InputBox("1- Consulta pelo código | 2- Consulta pelo nome (0 sai)")
        
        If Trim(n) = "" Then
            MsgBox "Consulta não efetuada!"
            'Exit Sub
        ElseIf n = 0 Then
            'MsgBox "Consulta não efetuada!"
            Exit Sub
        ElseIf n = 1 Then
            consultaEstoque
        ElseIf n = 2 Then
            consulNome
        Else
            MsgBox "Opção inválida! Tente novamente."
        End If
        
    Loop While n <> 0
End Sub

' Consulta o estoque pelo código
Sub consultaEstoque() '(codProduto As Long)
    Dim quant%, a#, i#, n As Variant, nome As String, totalLinhas#, cod%, val As Date, locali As Variant
    Dim pe As Worksheet
    
    Set pe = ThisWorkbook.Sheets("Estoque")
    totalLinhas = total()
    
    Do
        ' produto a ser verificado pelo código
        n = InputBox("Digite o código do produto: (0 sai)")
        
        If Not IsNumeric(n) Then
            'MsgBox "Consulta não efetuada!"
            'Exit Sub
        ElseIf n = 0 Then
            'MsgBox "Consulta não efetuada!"
            Exit Sub
        End If
        
        n = CInt(n)
        For i = 1 To totalLinhas ' aqui ele chama a variavel que recebou a função que conta a quantidade de célula
            If n = pe.Cells(i, "A").Value Then
                cod = pe.Cells(i, "A").Value
                quant = pe.Cells(i, "I").Value
                val = pe.Cells(i, "E").Value
                nome = pe.Cells(i, "B").Value
                locali = pe.Cells(i, "G").Value
                
                MsgBox "Código: " & cod & vbCrLf & "Produto: " & _
                nome & vbCrLf & "Validade: " & val & vbCrLf & "Quantidade: " & quant & vbCrLf & _
                "Prateleira: " & locali
                Exit Sub
            End If
        Next i
        
        MsgBox "Produto inválido. Tente novamente."
    
    Loop While n <> 0

End Sub

' Consulta o estoque pelo nome
Sub consulNome()

    Dim quant#, i#, n As Variant, totalLinhas#, cod#, val As Date, locali As Variant, cont#
    Dim pe As Worksheet
    Dim matriz() As Variant

    Set pe = ThisWorkbook.Sheets("Estoque")
    totalLinhas = total()

    Do
        n = InputBox("Qual nome do produto: (0 sai)")
        n = UCase(n)
        
        If n = 0 Then
            Exit Sub
        ElseIf n = "" Or IsNumeric(n) Then
            'MsgBox "Consulta nao efetuada!"
            'Exit Sub
        End If
        
        cont = 0
        ReDim matriz(1 To totalLinhas, 1 To 5)

        ' alimenta uma matriz com os produtos (as que vai ser exibida pela consulta)
        For i = 1 To totalLinhas
            If n = pe.Cells(i, "B").Value Then
                cont = cont + 1
                matriz(cont, 1) = pe.Cells(i, "A").Value ' Código
                matriz(cont, 2) = n ' Nome do produto
                matriz(cont, 3) = pe.Cells(i, "E").Value ' Validade
                matriz(cont, 4) = pe.Cells(i, "I").Value ' Quantidade
                matriz(cont, 5) = pe.Cells(i, "G").Value ' Prateleira
            End If
        Next i
        
        If cont = 0 Then
            MsgBox "Produto inválido. Tente novamente."
        Else
            Dim consulta As String
            consulta = "Resultados da consulta:" & vbCrLf
            ' joga em uma variavel o resultado da matriz
            ' aqui que vai fazer exibir se tiver mais do mesmo produto cadastrado (mesmos nomes)
            For i = 1 To cont
                consulta = consulta & vbCrLf & "Código: " & matriz(i, 1) & vbCrLf & _
                                 "Produto: " & matriz(i, 2) & vbCrLf & _
                                 "Validade: " & matriz(i, 3) & vbCrLf & _
                                 "Quantidade: " & matriz(i, 4) & vbCrLf & _
                                 "Prateleira: " & matriz(i, 5) & vbCrLf & vbCrLf
            Next i
            
            MsgBox consulta
            Exit Sub
        End If
    Loop

End Sub

Function total() As Double
    Dim j%
    Dim planilhaEstoque As Worksheet
    
    Set planilhaEstoque = ThisWorkbook.Sheets("Estoque")
    
    j = 1
    Do While planilhaEstoque.Cells(j, "A").Value <> ""
    j = j + 1
    Loop
    total = j - 1
End Function

