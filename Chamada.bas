Attribute VB_Name = "Chamada"
Sub principal()
    'Dim i As Integer
    Dim qtd As Integer
    Dim lin As Integer
    Dim lista() As String
                
    'Conta quantos alunos a turma tem
    lin = 1
    While Cells(lin, 1) <> ""
        lin = lin + 1
    Wend
    
    qtd = lin - 1
    ReDim lista(1 To 10)
    
    'Armazena todos os nomes em um array
    lin = 1
    While lin <= qtd
       lista(lin) = Cells(lin, 1)
       lin = lin + 1
    Wend
    
    Janela.quadroInterno.Caption = "Inteligência Artificial"
    
    i = 1
    While i <= qtd
        Janela.nomeAluno = lista(i)
        Janela.Show
        i = i + 1
    Wend
    
End Sub
