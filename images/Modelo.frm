VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modelo 
   Caption         =   "Aula"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10710
   OleObjectBlob   =   "Modelo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botaoAusente_Click()
    
    End
    
End Sub

Private Sub botaoPresente_Click()
    
    Dim col As Integer
    
    col = 6
    While Worksheets("Planilha").Cells(10, col) <> ""
        col = col + 1
    Wend
    
    If qtd2 Then
        Range(Worksheets("Planilha").Cells(9, col), Worksheets("Planilha").Cells(9, col + 1)).MergeCells = True
        Range(Worksheets("Planilha").Cells(9, col), Worksheets("Planilha").Cells(9, col + 1)).HorizontalAlignment = xlCenter
        Worksheets("Planilha").Cells(9, col) = CInt(mes)
        Worksheets("Planilha").Cells(10, col) = CInt(dia)
        Worksheets("Planilha").Cells(10, col + 1) = CInt(dia)
        
        Call preencherFrequencia(2)
    ElseIf qtd4 Then
        Range(Worksheets("Planilha").Cells(9, col), Worksheets("Planilha").Cells(9, col + 3)).MergeCells = True
        Range(Worksheets("Planilha").Cells(9, col), Worksheets("Planilha").Cells(9, col + 3)).HorizontalAlignment = xlCenter
        Worksheets("Planilha").Cells(9, col) = CInt(mes)
        Worksheets("Planilha").Cells(10, col) = CInt(dia)
        Worksheets("Planilha").Cells(10, col + 1) = CInt(dia)
        Worksheets("Planilha").Cells(10, col + 2) = CInt(dia)
        Worksheets("Planilha").Cells(10, col + 3) = CInt(dia)
        
        Call preencherFrequencia(4)
    End If
    
    Unload Me
    
End Sub

Private Sub CommandButton1_Click()
    Dim janela As FileDialog
    Dim arquivoEscolhido As Integer
    
    Set janela = Application.FileDialog(msoFileDialogFilePicker)
    arquivoEscolhido = janela.Show
    
    If arquivoEscolhido <> -1 Then
        'didn't choose anything (clicked on CANCEL)
        MsgBox "Você não selecionou nenhum arquivo"
    Else
        'display name and path of file chosen
        MsgBox janela.SelectedItems(1)
    End If
    
    
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)
    
    If closemode = vbFormControlMenu Then
        End
    End If
    
End Sub
