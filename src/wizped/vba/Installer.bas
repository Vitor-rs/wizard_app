Attribute VB_Name = "WizpedInstaller"
Option Explicit

' ==========================================================================================
' WIZPED - INSTALADOR DE SISTEMA (FINAL STABLE)
' ==========================================================================================
' Módulo DESCARTÁVEL.
'
' Instruções:
' 1. Importe este arquivo.
' 2. Execute a macro "InstalarCompleto".
' ==========================================================================================

Public Sub InstalarCompleto()
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    If Err.Number <> 0 Then
        MsgBox "Erro de permissão! Habilite 'Confiar no acesso ao modelo de objeto do projeto do VBA'.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 1. Criar/Reparar o Formulário
    Dim frmComp As Object
    Set frmComp = GetOrCreateComponent(vbProj, "frmWizpedEditor", 3)
    
    BuildFormControls frmComp
    InjectFormCode frmComp.CodeModule
    
    ' 2. Criar o Launcher
    Dim modComp As Object
    Set modComp = GetOrCreateComponent(vbProj, "modWizped", 1)
    InjectLauncherCode modComp.CodeModule
    
    MsgBox "Instalação Corrigida!", vbInformation
End Sub

' --- GESTÃO DE COMPONENTES ---

Private Function GetOrCreateComponent(vbProj As Object, name As String, compType As Integer) As Object
    Dim comp As Object
    On Error Resume Next
    Set comp = vbProj.VBComponents(name)
    On Error GoTo 0
    
    If comp Is Nothing Then
        Set comp = vbProj.VBComponents.Add(compType)
        comp.Name = name
    End If
    Set GetOrCreateComponent = comp
End Function

' --- CONSTRUÇÃO DO FORMULÁRIO ---

Private Sub BuildFormControls(vbComp As Object)
    Dim frm As Object
    Set frm = vbComp.Designer
    
    vbComp.Properties("Caption") = "Editor Wizped (Stable)"
    vbComp.Properties("Width") = 420
    vbComp.Properties("Height") = 340
    
    EnsureControl frm, "lstProdutos", "Forms.ListBox.1", 10, 10, 390, 150
    With frm.Controls("lstProdutos")
        .ColumnCount = 4
        .ColumnHeads = True
        .ColumnWidths = "80;150;60;50"
    End With
    
    EnsureControl frm, "lblSKU", "Forms.Label.1", 170, 10, 60, 18, "SKU:"
    EnsureControl frm, "txtSKU", "Forms.TextBox.1", 170, 70, 100, 18
    
    EnsureControl frm, "lblNome", "Forms.Label.1", 200, 10, 60, 18, "Nome:"
    EnsureControl frm, "txtNome", "Forms.TextBox.1", 200, 70, 200, 18
    
    EnsureControl frm, "lblPreco", "Forms.Label.1", 230, 10, 60, 18, "Preço:"
    EnsureControl frm, "txtPreco", "Forms.TextBox.1", 230, 70, 80, 18
    
    EnsureControl frm, "lblEstoque", "Forms.Label.1", 230, 160, 50, 18, "Estoque:"
    EnsureControl frm, "txtEstoque", "Forms.TextBox.1", 230, 210, 60, 18
    
    EnsureControl frm, "btnNovo", "Forms.CommandButton.1", 270, 10, 80, 24, "Novo"
    EnsureControl frm, "btnSalvar", "Forms.CommandButton.1", 270, 100, 90, 24, "Salvar"
    EnsureControl frm, "btnExcluir", "Forms.CommandButton.1", 270, 300, 90, 24, "Excluir"
End Sub

Private Sub EnsureControl(frm As Object, name As String, classID As String, t As Double, l As Double, w As Double, h As Double, Optional caption As String = "")
    Dim ctl As Object
    On Error Resume Next
    Set ctl = frm.Controls(name)
    On Error GoTo 0
    If ctl Is Nothing Then Set ctl = frm.Controls.Add(classID, name)
    ctl.Top = t: ctl.Left = l: ctl.Width = w: ctl.Height = h
    If caption <> "" Then ctl.Caption = caption
End Sub

' --- INJEÇÃO DE CÓDIGO (FORM) ---

Private Sub InjectFormCode(modCode As Object)
    modCode.DeleteLines 1, modCode.CountOfLines
    Dim code As String
    code = "Option Explicit" & vbCrLf & vbCrLf
    
    ' ----> LÓGICA DE EXECUÇÃO ABSOLUTA <----
    
    code = code & "Private Sub RunPython(args As String)" & vbCrLf
    
    ' Uso de FSO para resolver caminho absoluto
    code = code & "    Dim fso As Object" & vbCrLf
    code = code & "    Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    
    ' Tenta localizar o batch relativo ao workbook
    code = code & "    Dim batRel As String, batAbs As String" & vbCrLf
    code = code & "    ' Se workbook esta em src/wizped, o batch esta em src/debug_launcher.bat (sobe um nivel)" & vbCrLf
    code = code & "    batRel = fso.BuildPath(ThisWorkbook.Path, ""..\debug_launcher.bat"")" & vbCrLf
    
    ' Pega o caminho canonico (resolve o ..)
    code = code & "    batAbs = fso.GetAbsolutePathName(batRel)" & vbCrLf
    
    code = code & "    ' CHECK DE EXISTENCIA" & vbCrLf
    code = code & "    If Not fso.FileExists(batAbs) Then" & vbCrLf
    code = code & "         MsgBox ""ERRO CRITICO: Arquivo de boot nao encontrado!"" & vbCrLf & ""Procurado em: "" & batAbs, vbCritical" & vbCrLf
    code = code & "         Exit Sub" & vbCrLf
    code = code & "    End If" & vbCrLf
    
    ' EXECUÇÃO DIRETA (Evita cmd /c quote hell)
    code = code & "    Dim wsh As Object" & vbCrLf
    code = code & "    Set wsh = CreateObject(""WScript.Shell"")" & vbCrLf
    
    ' Aspas triplas para garantir que o path do arquivo tenha aspas no comando
    ' Run """C:\Path\To\File.bat"" arg1 arg2"
    code = code & "    Dim cmd As String" & vbCrLf
    code = code & "    cmd = """""""" & batAbs & """""""" & "" "" & args" & vbCrLf
    
    code = code & "    ' Executa e espera (0 = Hide)" & vbCrLf
    code = code & "    wsh.Run cmd, 0, True" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    code = code & "Private Sub UserForm_Initialize()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Dim tbl As ListObject" & vbCrLf
    code = code & "    Set tbl = ThisWorkbook.Sheets(""produtos"").ListObjects(""tbl_produtos"")" & vbCrLf
    code = code & "    If Not tbl Is Nothing Then" & vbCrLf
    code = code & "        lstProdutos.RowSource = tbl.DataBodyRange.Address(External:=True)" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    code = code & "Private Sub lstProdutos_Click()" & vbCrLf
    code = code & "    Dim i As Long: i = lstProdutos.ListIndex" & vbCrLf
    code = code & "    If i = -1 Then Exit Sub" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    txtSKU.Value = lstProdutos.List(i, 0)" & vbCrLf
    code = code & "    txtNome.Value = lstProdutos.List(i, 1)" & vbCrLf
    code = code & "    txtPreco.Value = lstProdutos.List(i, 2)" & vbCrLf
    code = code & "    txtEstoque.Value = lstProdutos.List(i, 3)" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    code = code & "Private Sub btnNovo_Click()" & vbCrLf
    code = code & "    txtSKU.Value = """": txtNome.Value = """"" & vbCrLf
    code = code & "    txtPreco.Value = """": txtEstoque.Value = """"" & vbCrLf
    code = code & "    lstProdutos.ListIndex = -1" & vbCrLf
    code = code & "    On Error Resume Next: txtSKU.SetFocus" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' BOTÃO SALVAR AGORA RECARREGA A LISTA
    code = code & "Private Sub btnSalvar_Click()" & vbCrLf
    code = code & "    Dim p As String" & vbCrLf
    code = code & "    p = Replace(txtPreco.Value, "","", ""."")" & vbCrLf
    code = code & "    Me.Repaint" & vbCrLf
    code = code & "    Dim args As String" & vbCrLf
    ' Usando chr(34) para facilitar aspas
    code = code & "    args = ""save --sku "" & Chr(34) & txtSKU.Value & Chr(34) & "" --nome "" & Chr(34) & txtNome.Value & Chr(34) & "" --preco "" & p & "" --estoque "" & txtEstoque.Value" & vbCrLf
    code = code & "    RunPython args" & vbCrLf
    code = code & "    UserForm_Initialize" & vbCrLf
    code = code & "    MsgBox ""Salvo!"", vbInformation" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' BOTÃO EXCLUIR - CORRIGIDO SINTAXE DE ASPAS
    code = code & "Private Sub btnExcluir_Click()" & vbCrLf
    code = code & "    If MsgBox(""Excluir produto?"", vbYesNo) = vbYes Then" & vbCrLf
    code = code & "        Dim args As String" & vbCrLf
    ' Correção aqui usando Chr(34) também
    code = code & "        args = ""delete --sku "" & Chr(34) & txtSKU.Value & Chr(34)" & vbCrLf
    code = code & "        RunPython args" & vbCrLf
    code = code & "        btnNovo_Click" & vbCrLf
    code = code & "        UserForm_Initialize" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf

    modCode.AddFromString code
End Sub

' --- INJEÇÃO DE CÓDIGO (LAUNCHER) ---

Private Sub InjectLauncherCode(modCode As Object)
    modCode.DeleteLines 1, modCode.CountOfLines
    Dim code As String
    code = "Option Explicit" & vbCrLf & vbCrLf
    code = code & "Public Sub AbrirWizped()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    If ThisWorkbook.Sheets(""produtos"").ListObjects(""tbl_produtos"") Is Nothing Then" & vbCrLf
    code = code & "        MsgBox ""Erro: Tabela 'tbl_produtos' nao encontrada."", vbCritical" & vbCrLf
    code = code & "        Exit Sub" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    frmWizpedEditor.Show vbModeless" & vbCrLf
    code = code & "End Sub"
    modCode.AddFromString code
End Sub
