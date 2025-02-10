Attribute VB_Name = "Retirada"
Option Explicit

Sub retirada() ' � um "menu" das op��es de retirada e algumas valida��es
    Dim n As Variant
    
    Do
        n = InputBox("1- Venda | 2- Remo��o | (0 sai)")
        
        If Not IsNumeric(n) Then
            MsgBox "Insira uma op��o valida"
            retirada
            Exit Sub
        End If
        
        n = CInt(n)

        If n = 0 Then
            'MsgBox "Consulta n�o efetuada!"
            Exit Sub
        ElseIf n = 1 Then
            venda
            Exit Sub
        ElseIf n = 2 Then
            remocao
            Exit Sub
        Else
            MsgBox "Op��o inv�lida! Tente novamente."
        End If
        
    Loop While n <> 0
End Sub


Sub venda() ' por meio do c�digo faz uma movimenta��o na quantidade em estoque. e tambem faz valida��es e tratamento de erro

    Dim pm As Worksheet
    Dim pe As Worksheet
    Dim i%, cod As Variant, quant%, totalEstoque%, data%, busca As Boolean

    Set pm = ThisWorkbook.Sheets("Movimenta��o")
    Set pe = ThisWorkbook.Sheets("Estoque")
    
    totalEstoque = pe.Cells(pe.Rows.Count, "I").End(xlUp).Row
    busca = False
    
    cod = InputBox("Informe o c�digo do produto vendido.")
    
    If Not IsNumeric(cod) Then 'Or cod = 0 Then
        MsgBox "Insira um c�digo valido!"
        venda
        Exit Sub
    End If
    
    cod = CInt(cod)

    ' Recusa produto vencido de ser vendido
    For i = 2 To totalEstoque
         If pe.Cells(i, "A") = cod Then
            busca = True ' Marca como encontrado
            data = DateDiff("d", Date, CDate(pe.Cells(i, "E").Value))
        
            If data <= 0 Then
                MsgBox "Produto vencido, n�o aceita venda!"
                Exit Sub
            End If
        End If
    Next i
    
    If Not busca Then
        MsgBox "C�digo n�o cadastrado!"
        Exit Sub
    End If
     If Not IsNumeric(cod) Then 'Or cod = 0 Then
        MsgBox "Insira um c�digo valido!"
        venda
        Exit Sub
    End If
    
    
    quant = InputBox("Quantos foram vendidos?")
    
    For i = 2 To totalEstoque
        If pe.Cells(i, "A") = cod Then
            busca = True ' Marca como encontrado
            
            
            ' faz a movimenta��o entre planilhas
            If pe.Cells(i, "A") = cod And pe.Cells(i, "I") < quant Then
                MsgBox "Quantidade solicitada n�o dispon�vel!" & vbCrLf & vbCrLf & "Quantidade em estoque: " & pe.Cells(i, "I")
                Exit Sub
            ElseIf quant > 0 And pe.Cells(i, "A") = cod Then
                pm.Range("A2:E2").Insert
                pm.Range("A2") = pe.Cells(i, "A")
                pm.Range("B2") = pe.Cells(i, "B")
                pm.Range("C2") = pe.Cells(i, "E")
                pm.Range("D2") = Date
                pm.Range("E2") = quant
                pe.Cells(i, "I") = pe.Cells(i, "I") - quant
                
                    If pe.Cells(i, "I") = 0 Then
                        pm.Range("E2") = "Todo o estoque vendido!"
                        pe.Cells(i, "A").EntireRow.Delete
                        Exit Sub
                    End If
                Exit Sub
            End If
        End If
    Next i
    
End Sub

Sub remocao() ' Faz a remo��o do produto
    
    Dim pm As Worksheet
    Dim pe As Worksheet
    Dim i%, cod As Variant, quant%, totalEstoque%, data%, busca As Boolean

    Set pm = ThisWorkbook.Sheets("Movimenta��o")
    Set pe = ThisWorkbook.Sheets("Estoque")
    
    totalEstoque = pe.Cells(pe.Rows.Count, "I").End(xlUp).Row
    
    cod = InputBox("Informe o c�digo do produto para sa�da.")
    
    If Not IsNumeric(cod) Then
        MsgBox "O c�digo informado � inv�lido. Tente novamente."
        remocao
        Exit Sub
    End If
    
    cod = CInt(cod)
    For i = 2 To totalEstoque
        If pe.Cells(i, "A") = cod Then
            pm.Range("A2:E2").Insert
            pm.Range("A2:E2").Interior.Color = RGB(197, 217, 241) 'azul claro
            pm.Range("A2") = pe.Cells(i, "A")
            pm.Range("B2") = pe.Cells(i, "B")
            pm.Range("C2") = pe.Cells(i, "E")
            pm.Range("D2") = Date
            pm.Range("E2") = pe.Cells(i, "I")
            pe.Cells(i, "A").EntireRow.Delete
            busca = True ' Marca como encontrado
            Exit For ' Sai do loop assim que encontrar
        End If
    Next i
    
    If Not busca Then
        MsgBox "O c�digo informado n�o foi encontrado. Verifique e tente novamente."
        retirada
        Exit Sub
    End If
End Sub
