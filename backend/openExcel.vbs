Option Explicit

Dim excelApp
Set excelApp = CreateObject("Excel.Application")

' Definir caminho para o arquivo Excel
Dim filePath, sheetName, cellAddress
filePath = "C:\Users\DEYVSON FELIPE\Desktop\busca-arquivos\ARQUIVO.xlsm" ' Substitua pelo caminho correto para o seu arquivo Excel
sheetName = WScript.Arguments(0)
cellAddress = WScript.Arguments(1)

' Abrir o arquivo Excel
Dim workbook
Set workbook = excelApp.Workbooks.Open(filePath)
excelApp.Visible = True

' Executar a macro VBA para destacar a linha inteira
On Error Resume Next
excelApp.Run "GoToSheetAndCell", sheetName, cellAddress
If Err.Number <> 0 Then
    WScript.Echo "Erro ao executar a macro VBA: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

' Aguardar um momento para o destaque ser visível
WScript.Sleep 2000 ' Aguarda 2 segundos (ajuste conforme necessário)

' Fechar o arquivo Excel sem salvar alterações
workbook.Close False
excelApp.Quit

' Liberar objetos da memória
Set workbook = Nothing
Set excelApp = Nothing
