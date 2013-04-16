VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frequencia 
   Caption         =   "Frequência"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13425
   OleObjectBlob   =   "Frequencia.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frequencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botaoAusente_Click()
    
    Dim col As Integer
    
    col = 6
    While Worksheets("Planilha").Cells(linhaAluno, col) <> ""
        col = col + 1
    Wend
    
    If qtd = 2 Then
        Worksheets("Planilha").Cells(linhaAluno, col) = "F"
        Worksheets("Planilha").Cells(linhaAluno, col + 1) = "F"
    ElseIf qtd = 4 Then
        Worksheets("Planilha").Cells(linhaAluno, col) = "F"
        Worksheets("Planilha").Cells(linhaAluno, col + 1) = "F"
        Worksheets("Planilha").Cells(linhaAluno, col + 2) = "F"
        Worksheets("Planilha").Cells(linhaAluno, col + 3) = "F"
    End If
    
    Unload Me
    
End Sub

Private Sub botaoPresente_Click()

    Dim col As Integer
    
    col = 6
    While Worksheets("Planilha").Cells(linhaAluno, col) <> ""
        col = col + 1
    Wend
    
    If qtd = 2 Then
        Worksheets("Planilha").Cells(linhaAluno, col) = "P"
        Worksheets("Planilha").Cells(linhaAluno, col + 1) = "P"
    ElseIf qtd = 4 Then
        Worksheets("Planilha").Cells(linhaAluno, col) = "P"
        Worksheets("Planilha").Cells(linhaAluno, col + 1) = "P"
        Worksheets("Planilha").Cells(linhaAluno, col + 2) = "P"
        Worksheets("Planilha").Cells(linhaAluno, col + 3) = "P"
    End If
    
    Unload Me
    
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)
    
    If closemode = vbFormControlMenu Then
        
        End
        
        'MsgBox "Sorry you must use Cancel or Exit Button"
        'Cancel = True
        
    End If
    
End Sub
