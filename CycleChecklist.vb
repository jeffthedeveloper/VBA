Sub Checklist_Cíclico()

' Declaração de variáveis
Dim ws As Worksheet
Dim txtOS As String
Dim chkEtapa As CheckBox
Dim btnCriarPasta As Button
Dim btnBaixarImagens As Button
Dim btnAbrirLink As Button
Dim btnCriarCheckin As Button
Dim btnAnexarImagens As Button
Dim btnCriarRelatorio As Button

' Obter a planilha ativa
Set ws = ThisWorkbook.ActiveSheet

' Obter os elementos da interface
Set txtOS = ws.Shapes("txtOS").OLEFormat.Object
Set chkEtapa = ws.Shapes("chkEtapa")
Set btnCriarPasta = ws.Shapes("btnCriarPasta").OLEFormat.Object
Set btnBaixarImagens = ws.Shapes("btnBaixarImagens").OLEFormat.Object
Set btnAbrirLink = ws.Shapes("btnAbrirLink").OLEFormat.Object
Set btnCriarCheckin = ws.Shapes("btnCriarCheckin").OLEFormat.Object
Set btnAnexarImagens = ws.Shapes("btnAnexarImagens").OLEFormat.Object
Set btnCriarRelatorio = ws.Shapes("btnCriarRelatorio").OLEFormat.Object

' Iniciar o processo
If txtOS.Text = "" Then
    MsgBox "Digite o número da OS (4 dígitos)"
    Exit Sub
End If

' Executar as etapas do processo
' ... (substituir por código específico para cada etapa)

' Marcar todas as tarefas como concluídas
For Each chkEtapa In ws.Shapes
    If chkEtapa.Type = msoShapeCheckbox Then
        chkEtapa.Value = True
    End If
Next

' Reiniciar o processo
MsgBox "Processo finalizado!"
txtOS.Text = ""

End Sub
