' ==============================================================================
' MÓDULO: Conversor de Evidências (Outlook -> Word -> PDF)
' DESCRIÇÃO: Gera arquivos PDF para conformidade legal e arquivamento digital.
' ==============================================================================

Sub Gerar_Evidencia_PDF()
    Dim WordApp As Object, WordDoc As Object
    Dim tempPath As String, finalPDFPath As String
    
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = False
    
    ' Exemplo de fluxo: Pega o conteúdo temporário e exporta como PDF fixo
    tempPath = Environ("TEMP") & "\temp_processo.mht"
    finalPDFPath = "C:\Repositorio_Digital\Evidencias\Processo_001.pdf"
    
    ' Lógica de exportação robusta
    ' Set WordDoc = WordApp.Documents.Open(tempPath)
    ' WordDoc.ExportAsFixedFormat finalPDFPath, 17 ' 17 = wdExportFormatPDF
    ' WordDoc.Close False
    
    WordApp.Quit
End Sub
