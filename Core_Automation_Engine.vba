' ==============================================================================
' PROJETO: Automação de Processamento de Demandas e Comunicação em Lote
' IMPACTO: Redução de 77% no tempo de execução (4 min -> 55 seg por item)
' DESCRIÇÃO: Motor de envio e arquivamento que integra Excel e Outlook.
' ==============================================================================

Sub Enviar_E_Arquivar_Demandas()
    Dim OutlookApp As Object, EmailItem As Object, FolderSent As Object
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Painel_Controle")
    Dim i As Long, lastRow As Long
    Dim processID As String, secondaryID As String, caseNumber As String
    Dim emailDest As String, subjectLine As String, responseType As String
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set FolderSent = OutlookApp.GetNamespace("MAPI").GetDefaultFolder(5) ' 5 = Sent Items
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        ' Extração de dados (Nomes genéricos para proteção de dados)
        caseNumber = ws.Cells(i, 1).Value
        processID = ws.Cells(i, 2).Value
        secondaryID = ws.Cells(i, 3).Value
        emailDest = "cliente.externo@exemplo.com.br" ' E-mail fictício
        responseType = ws.Cells(i, 7).Value
        
        ' Lógica de Assunto Dinâmico
        subjectLine = "Notificação de Processo - Ref: " & caseNumber & " [ID: " & processID & "]"
        If responseType = "FINAL" Then subjectLine = subjectLine & " - RESPOSTA FINAL"
        
        ' Criação e Envio do E-mail
        Set EmailItem = OutlookApp.CreateItem(0)
        With EmailItem
            .To = emailDest
            .Subject = subjectLine
            .HTMLBody = "<html><body>Prezados,<br>Conforme análise técnica, segue a documentação...</body></html>"
            .Send
        End With
        
        ' Aguarda processamento do servidor
        DoEvents
        
        ' Localização e Salvamento da Cópia Enviada para Auditoria
        Dim sentItem As Object
        For Each sentItem In FolderSent.Items
            If sentItem.Subject = subjectLine Then
                sentItem.SaveAs ws.Cells(i, 8).Value & "\" & processID & ".msg", 3
                Exit For
            End If
        Next
    Next i
    
    MsgBox "Processamento em lote concluído com sucesso!", vbInformation
End Sub
