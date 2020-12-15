Attribute VB_Name = "GeradorHex"
Sub obterNumBinario()
    Dim numBin As String
    Dim numHex As String
    
    Dim quantBits As Integer
    
    quantBits = 0

'    Percorre o desenho
    For col = 5 To 20
        For lin = 8 To 23
            
            If Cells(lin, col).Value = 1 Then
                numBin = numBin & 1
            Else
                numBin = numBin & 0
            End If
            
            quantBits = quantBits + 1
                
            If quantBits = 4 Then
'                Manda o valor Binario Obtido para conversao
                Cells(1, 2).Value = numBin
                numHex = numHex & Cells(1, 1).Value

'                Reseta variaveis
                Cells(1, 2).Value = ""
                numBin = ""
                quantBits = 0
            End If
        Next lin
        Cells(25, col).Value = numHex
        numHex = ""
    Next col
    
End Sub

Sub limparMatriz()
    resposta = MsgBox("Deseja limpar a matriz?" & Chr(13) & "Essa ação não pode ser desfeita!", vbYesNo, "Limpar matriz")
    If resposta = vbYes Then
'        Limpa a matriz
        Range("E8:T23").Select
        Selection.ClearContents
        
'        Limpa os valores Hexadecimais
        Range("E25:T25").Select
        Selection.ClearContents
        
        Range("E8").Select
    End If
End Sub
