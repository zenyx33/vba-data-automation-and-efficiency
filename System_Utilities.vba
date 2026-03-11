' ==============================================================================
' MÓDULO: Utilitários de Sistema e Governança de Dados
' ==============================================================================

Sub Gerenciar_Diretorios()
    Dim basePath As String: basePath = "C:\Operacoes_Dados\"
    Dim folderName As String
    ' Lógica para garantir que a estrutura de pastas exista antes do processamento
    If Dir(basePath, vbDirectory) = "" Then MkDir basePath
End Sub

Sub Limpar_Ambiente_Trabalho()
    ' Reseta a planilha para o próximo ciclo de processamento
    Range("A2:H100").ClearContents
End Sub
