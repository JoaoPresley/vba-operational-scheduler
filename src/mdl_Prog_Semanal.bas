Attribute VB_Name = "mdl_Prog_Semanal"
Option Explicit
Dim ws_Dst As Worksheet 'Planilha semanal
Dim dictColunas As Object 'Nome dos tecnicos da planilha SEMANAL para Coluna

Sub Updated_data()
    Application.ScreenUpdating = False 'Desativa a atualização da tela
    'Em caso de erro apresenta o erro e para o código
    On Error GoTo ErrHandler
    
    'Define variaveis
    Dim DayStart As Date 'Dia do inicio da semana
    Dim DayEnd As Date 'Dia do fim da semana
    Dim Texto As String 'Texto que será escrito no topo de cada "calendário"
    Dim Sufixo As String 'Data que vai no final de cada texto
    Dim LabelManutenção As Range 'Célula do topo do calendário MANUTENÇÂO
    Dim LabelOperação As Range 'Célula do topo do calendário OPERAÇÃO
    Dim LabelTreslagoas As Range 'Célula do topo do calendário TRES LAGOAS
    
    '------------------
    'Atribuições das variaveis
    Set ws_Dst = ThisWorkbook.Worksheets("SEMANAL")
    Set LabelManutenção = ws_Dst.Cells(2, 2)
    Set LabelOperação = ws_Dst.Cells(2, "K")
    Set LabelTreslagoas = ws_Dst.Cells(2, "T")
    '   Atribui os dias iniciais e finais
    DayStart = ws_Dst.Cells(5, "C").Value
    DayEnd = ws_Dst.Cells(15, "C").Value
    '   Define o sufixo que é mesmo para todos os textos
    Sufixo = Format(DayStart, "dd/mm/yyyy") & " A " & Format(DayEnd, "dd/mm/yyyy")
    
    '------------------
    'Escreve o dia no topo dos calendários
    LabelManutenção = "PROGRAMAÇÃO SEMANAL - MANUTENÇÃO - " & Sufixo
    LabelOperação = "PROGRAMAÇÃO SEMANAL - OPERAÇÃO - " & Sufixo
    LabelTreslagoas = "PROGRAMAÇÃO SEMANAL - TRÊS LAGOAS - " & Sufixo
    
ExitHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

Sub fill_work(setor As String)
    'Desativa atualização imediata da tela e cria caminho para erro
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler
    
    
    Dim lastRowSrc As Long
    Dim ws_Src As Worksheet
    Set ws_Src = ThisWorkbook.Worksheets(setor)
    Set ws_Dst = ThisWorkbook.Worksheets("SEMANAL")
    
    '----------------------------
    'Pega a ultima linha relevante da ws_Src
    ' usar a coluna H (tipo do serviço) para identificar fim dos registros; caso vazio, varrer até a última linha usada
    For lastRowSrc = 6 To ws_Src.Rows.Count
        If ws_Src.Cells(lastRowSrc, "H") = "Sobr." Then
            Exit For
        End If
    Next lastRowSrc
    lastRowSrc = lastRowSrc - 1
    
    If lastRowSrc < 6 Then lastRowSrc = 6
    
    '----------------------------
    'Limpa a planilha
    Call LimpaPlanilha(setor, ws_Dst)
    
    '------------------------
    'Pega os serviços de cada dia
    Dim Toms(1 To 2) As String 'Array com o nome dos tecnicos
    Dim Serviço As String 'Serviço que técnico fará
    
    
    'Semana vai da coluna M (13) até X (24) de dois em dois
    Dim col As Integer
    Dim lin As Integer
    Dim col_insert As Integer 'Valor da coluna que colocará o serviço
     'Dim lin_insert As Integer 'Valor da linha que colocará o serviço
    'Verifica cada coluna da manhã
    For col = 13 To 24
        For lin = 6 To lastRowSrc
            'Se a célula tiver valor verifica o nome
            If Not IsEmpty(ws_Src.Cells(lin, col)) Then
                Toms(1) = ws_Src.Cells(lin, "j").Value 'Primeiro TOM
                Toms(2) = ws_Src.Cells(lin, "k").Value 'Segundo TOM
                Serviço = ws_Src.Cells(lin, "l").Value 'Serviço dos TOMS
                
                'Prenche o serviço na planilha SEMANAL
                'Se os nomes forem existentes ele preenche
                If Not Toms(1) = "" Then
                    col_insert = nomeToColuna(Toms(1))
                    'Se o col_insert for validado insere o serviço
                    Debug.Print "Tom 1: " & Toms(1) & " col: " & (col - 8)
                    If Not col_insert = 0 Then ws_Dst.Cells(col - 8, col_insert) = Serviço
                End If
                If Not Toms(2) = "" Then
                    col_insert = nomeToColuna(Toms(2))
                    'Se o col_insert for validado insere o serviço
                    Debug.Print "Tom 2 col: " & (col - 8)
                    If Not col_insert = 0 Then ws_Dst.Cells(col - 8, col_insert) = Serviço
                End If
                
            End If
        Next lin
    Next col
    
    
ExitHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical
    Resume ExitHandler
End Sub
Private Sub CarregarMapeamento()
    Set dictColunas = CreateObject("Scripting.Dictionary")
    Dim ws_Dst As Worksheet
    Set ws_Dst = ThisWorkbook.Worksheets("SEMANAL")
    
    Dim c As Long
    Dim TOM As String
    ' Lê a linha 4 (onde estão os nomes) e guarda a coluna de cada um
    For c = 5 To 30 ' Ajuste o range conforme necessário
        TOM = Trim(ws_Dst.Cells(4, c).Value)
        If TOM <> "" Then
            'Verifica se o nome não é Técnico, que é apenas uma label e não um nome
            If TOM <> "Técnico" Then
                ' Adiciona ao dicionário: Nome -> Coluna
                dictColunas(Trim(ws_Dst.Cells(4, c).Value)) = c
            End If
        End If
    Next c
End Sub
Private Sub LimpaPlanilha(setor As String, planilha As Worksheet)
    Dim i As Integer
    Select Case setor
        Case "OCGR"
            planilha.Range("N5:R16").Value = "A definir"
        Case "MCGR"
            planilha.Range("E5:I16").Value = "A definir"
        Case "TLG"
            planilha.Range("W5:AA16").Value = "A definir"
    End Select
End Sub
Private Function nomeToColuna(nome As String) As Integer
    nome = Trim(nome)
    
    ' Se o dicionário ainda não foi carregado, carrega agora
    If dictColunas Is Nothing Then CarregarMapeamento
    
    ' Verifica se o nome existe
    If dictColunas.Exists(nome) Then
        nomeToColuna = dictColunas(nome)
    Else
        nomeToColuna = 0 ' Nome não encontrado
    End If
End Function


