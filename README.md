# QUEBRA_CSV_VBA
# Quebrar arquivo csv baseado em duas colunas e renomeando as mesmas para facilitar a localização

Sub SplitCSV()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cd_lojaCol As Range
    Dim REGIONALCol As Range
    Dim cd_loja As String
    Dim REGIONAL As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
   
    Set ws = ThisWorkbook.Sheets(1) ' Ajuste conforme necessário
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
   
    Set cd_lojaCol = ws.Range("B2:B" & lastRow)
    Set REGIONALCol = ws.Range("C2:C" & lastRow)
   
    Dim cell As Range
    For Each cell In REGIONALCol
        cd_loja = cell.Value
        REGIONAL = cell.Offset(0, -1).Value
        If Not dict.exists(cd_loja & "_" & REGIONAL) Then
            dict.Add cd_loja & "_" & REGIONAL, New Collection
        End If
        dict(cd_loja & "_" & REGIONAL).Add cell.EntireRow
    Next cell
   
    Dim key As Variant
    Dim newWS As Worksheet
    Dim i As Long
    For Each key In dict.keys
        Set newWS = ThisWorkbook.Sheets.Add
        newWS.Name = key
        ws.Rows(1).Copy newWS.Rows(1)
        i = 2
        For Each Row In dict(key)
            Row.Copy newWS.Cells(i, 1)
            i = i + 1
        Next Row
        newWS.Move
        ActiveWorkbook.SaveAs "Z:\Arquivo Pessoal - Equipe\Amanda\DIRETORIA_1\" & key & ".csv", xlCSV
        ActiveWorkbook.Close SaveChanges:=False
    Next key
End Sub
