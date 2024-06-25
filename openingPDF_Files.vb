Sub OpenPdfs()

'Declare variables
Dim filename As String
Dim i As Integer

'Loop through all files in the second sheet
For i = 1 To Sheets("Planilha2").Range("A1").End(xlDown).Row

'Get the filename from the cell
filename = Sheets("Planilha2").Cells(i, 1).Value

'Open the file
Shell "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe" & filename

Next i

End Sub
