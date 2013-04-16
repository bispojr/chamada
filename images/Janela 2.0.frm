VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Janela 
   Caption         =   "Frequência"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10905
   OleObjectBlob   =   "Janela 2.0.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Janela"
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
    
    Worksheets("Planilha").Cells(linhaAluno, col) = "F"
    
    Unload Me
    
End Sub

Private Sub botaoPresente_Click()

    Dim col As Integer
    
    col = 6
    While Worksheets("Planilha").Cells(linhaAluno, col) <> ""
        col = col + 1
    Wend
    
    Worksheets("Planilha").Cells(linhaAluno, col) = "P"
    
    Unload Me
    
End Sub

Private Sub sairChamada_Click()
    End
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, closemode As Integer)
    
    If closemode = vbFormControlMenu Then
        
        End
        
        'MsgBox "Sorry you must use Cancel or Exit Button"
        'Cancel = True
        
    End If
    
End Sub
