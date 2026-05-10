Attribute VB_Name = "mdl_Prog_Diária"
    Option Explicit
Public Const Ccol = 3, Dcol = 4, Ecol = 5, Fcol = 6, Gcol = 7
Dim dateTomorrow As Variant 'Dia da programação
Dim morningColSrc As Long, afternoonColSrc As Long 'Colunas da promação daquele dia, manhã e tarde respectivamente
Dim lastRowSrc As Long 'Ultima linha da planilha de programação
Dim Final_Tarde As Long


' Coloque este código dentro do objeto de planilha "OCGR - DIÁRIA"
Sub Prog_Diária(setor As String, Setor_Diaria As String)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wsSrc As Worksheet, wsDst As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(setor)
    Set wsDst = ThisWorkbook.Worksheets(Setor_Diaria)
    
    dateTomorrow = wsDst.Range("F2").Value ' =HOJE()+1
    
    
    
    ' --- limpar conteúdo antigo (Colunas C até G onde tem as OSs)
    ' --Encontra a região para apagar
    Dim Final_Manha As Long
    For Final_Manha = 1 To wsDst.Rows.Count
        If wsDst.Cells(Final_Manha, 4) = "Tarde" Then
            Exit For
        End If
    Next Final_Manha
    Final_Manha = Final_Manha - 1
    
    For Final_Tarde = 1 To wsDst.Rows.Count
        If wsDst.Cells(Final_Tarde, 3) = "Status" Then
            Exit For
        End If
    Next Final_Tarde
    Final_Tarde = Final_Tarde - 2
    
    ' --- limpar conteúdo antigo (colunas C, D, E, F e G a partir da linha 6) ---
    wsDst.Range(wsDst.Cells(6, Ccol), wsDst.Cells(Final_Manha, Gcol)).ClearContents
    wsDst.Range(wsDst.Cells(Final_Manha + 2, Ccol), wsDst.Cells(Final_Tarde, Gcol)).ClearContents
    
    ' ---Reseta para deixar apenas 2 linhas de manhã e 2 linhas de tarde
    ' Parte da manhã
    For Final_Manha = Final_Manha To 7 Step -1
        If Final_Manha = 7 Then
            Exit For
        End If
        wsDst.Rows(Final_Manha).Delete
        Final_Tarde = Final_Tarde - 1 'atualiza o final tarde
    Next Final_Manha
    'Parte da tarde
    For Final_Tarde = Final_Tarde To 10 Step -1
        If Final_Tarde = 10 Then
            Exit For
        End If
        wsDst.Rows(Final_Tarde).Delete
    Next Final_Tarde
    
    
    ' --- localizar a coluna da manhã/tarde na planilha OCGR ---
    ' colunas M:X correspondem a 13..24 (M=13,N=14,O=15,...X=24)
    
    morningColSrc = 0: afternoonColSrc = 0
    Dim c As Long
    For c = 13 To 24 Step 2
        ' valor da célula na linha 4 (coluna da manhã, célula mesclada à esquerda)
        If IsDate(wsSrc.Cells(4, c).Value) Then
            If CLng(wsSrc.Cells(4, c).Value) = CLng(dateTomorrow) Then
                morningColSrc = c
                afternoonColSrc = c + 1
                Exit For
            End If
        End If
    Next c
    
    If morningColSrc = 0 Then
        MsgBox "Data (" & dateTomorrow & ") não encontrada na linha 4 da planilha 'OCGR'. Verifique as datas.", vbExclamation
        GoTo ExitHandler
    End If
    
    ' --- preparar leitura dos registros (linhas a partir da 5) ---
    ' usar a coluna H (tipos de serviço) para identificar fim dos registros; caso vazio, varrer até a última linha usada
    For lastRowSrc = 6 To wsSrc.Rows.Count
        If wsSrc.Cells(lastRowSrc, "H") = "Folga" Then
            Exit For
        End If
    Next lastRowSrc
    lastRowSrc = lastRowSrc - 1
    
    If lastRowSrc < 6 Then lastRowSrc = 6
    
    ' listas de registros para manhã e tarde — cada item será um array: (techsCollection, OS, Servico)
    Dim morningEntries As Collection, afternoonEntries As Collection
    Set morningEntries = New Collection
    Set afternoonEntries = New Collection
    
    Dim r As Long
    For r = 6 To lastRowSrc
        Dim osM As Variant, osT As Variant
        osM = Trim(wsSrc.Cells(r, morningColSrc).Value)
        osT = Trim(wsSrc.Cells(r, afternoonColSrc).Value)
        
        Dim servico As String
        servico = Trim(wsSrc.Cells(r, "L").Value)
        
        ' pegar nomes técnicos (I,J,K) na ordem I,J,K (se estiverem preenchidos)
        Dim techs As New Collection
        Set techs = Nothing
        
        'If Trim(wsSrc.Cells(r, "I").Value) <> "" Then techs.Add Trim(wsSrc.Cells(r, "I").Value)
        If Trim(wsSrc.Cells(r, "J").Value) <> "" Then techs.Add Trim(wsSrc.Cells(r, "J").Value)
        If Trim(wsSrc.Cells(r, "K").Value) <> "" Then techs.Add Trim(wsSrc.Cells(r, "K").Value)
        
        ' apenas considerar se houver OS preenchida (conforme você confirmou)
        If osM <> "" Then
            ' para segurança: só adicionar se houver pelo menos 1 técnico
            If techs.Count > 0 Then
                Dim itemM As New Collection '(1 To 3) As Variant
                Set itemM = Nothing
                
                itemM.Add techs
                itemM.Add osM
                itemM.Add servico
                morningEntries.Add itemM
            End If
        End If
        
        If osT <> "" Then
            If techs.Count > 0 Then
                Dim itemT As New Collection
                Set itemT = Nothing
                
                itemT.Add techs
                itemT.Add osT
                itemT.Add servico
                afternoonEntries.Add itemT
            End If
        End If
    Next r
   
    'Preenche a quantidade de linhas na area da manhã e area da tarde
    Dim i As Long
    'Manhã
    If morningEntries.Count >= 2 Then
        For i = 2 To morningEntries.Count
               wsDst.Rows(Final_Manha).Insert
               Final_Manha = Final_Manha + 1 'Atualiza
               Final_Tarde = Final_Tarde + 1 'Atualiza
        Next i
    End If
    'Tarde
    If afternoonEntries.Count >= 2 Then
        For i = 2 To afternoonEntries.Count
               wsDst.Rows(Final_Tarde).Insert
               Final_Tarde = Final_Tarde + 1
        Next i
    End If
    
    ' --- Preencher MANHÃ começando na linha Final_Manhã - 1 ---
    Dim writeRow As Long
    writeRow = 6
    Dim rec As Long
    Dim coll As Collection
    For rec = 1 To morningEntries.Count
        Set coll = morningEntries(rec)(1)
        Dim osVal As Variant, servVal As String
        osVal = morningEntries(rec)(2)
        servVal = morningEntries(rec)(3)
        
        Select Case coll.Count
            Case 1
                wsDst.Cells(writeRow, Ccol).ClearContents '.Value = coll(1)
                wsDst.Cells(writeRow, Dcol).Value = coll(1) '.ClearContents
                wsDst.Cells(writeRow, Ecol).Value = osVal
                wsDst.Cells(writeRow, Fcol).Value = setor
                wsDst.Cells(writeRow, Gcol).Value = servVal
                writeRow = writeRow + 1
            Case 2
                wsDst.Cells(writeRow, Ccol).Value = coll(1)
                wsDst.Cells(writeRow, Dcol).Value = coll(2)
                wsDst.Cells(writeRow, Ecol).Value = osVal
                wsDst.Cells(writeRow, Fcol).Value = setor
                wsDst.Cells(writeRow, Gcol).Value = servVal
                writeRow = writeRow + 1
        End Select
    Next rec
    
    ' --- deixar uma linha em branco entre MANHÃ e TARDE ---
    Dim startAfternoonRow As Long
    'If writeRow = 6 Then startAfternoonRow = 12 Else startAfternoonRow = writeRow + 2
    startAfternoonRow = Final_Manha + 2
    
    ' --- Preencher TARDE começando em startAfternoonRow ---
    writeRow = startAfternoonRow
    For rec = 1 To afternoonEntries.Count
        Set coll = afternoonEntries(rec)(1)
        osVal = afternoonEntries(rec)(2)
        servVal = afternoonEntries(rec)(3)
        
        Select Case coll.Count
            Case 1
                wsDst.Cells(writeRow, Ccol).ClearContents
                wsDst.Cells(writeRow, Dcol).Value = coll(1)
                wsDst.Cells(writeRow, Ecol).Value = osVal
                wsDst.Cells(writeRow, Fcol).Value = setor
                wsDst.Cells(writeRow, Gcol).Value = servVal
                writeRow = writeRow + 1
            Case 2
                wsDst.Cells(writeRow, Ccol).Value = coll(1)
                wsDst.Cells(writeRow, Dcol).Value = coll(2)
                wsDst.Cells(writeRow, Ecol).Value = osVal
                wsDst.Cells(writeRow, Fcol).Value = setor
                wsDst.Cells(writeRow, Gcol).Value = servVal
                writeRow = writeRow + 1
        End Select
    Next rec
    
    ' Ajuste visual: remover seleção
    wsDst.Cells(1, 1).Select

ExitHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical
    Resume ExitHandler
End Sub
'Sub que preenche o Status do Técnico
Sub techs_state(base As String, Diária As String)
    'Pega a página que vai trabalhar------------------------
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(base)
    Set wsDst = ThisWorkbook.Worksheets(Diária)
    
    Dim start_vacation_line As Long 'Linha que tem os Techs de Folga/ Compensando/ Sexta + Leve
    Dim end_vacation_line As Long 'Linha que acaba os Techs de Folga/ Compensando/ Sexta + Leve
    start_vacation_line = lastRowSrc + 1
    
    'Varre as coluna H para encontrar a end_vacation_line
    For end_vacation_line = start_vacation_line To wsSrc.Rows.Count
        If (wsSrc.Cells(end_vacation_line, "H") = "Sobr.") Then
            Exit For
        End If
    Next end_vacation_line
    end_vacation_line = end_vacation_line - 1
    
    'Cria um dicionario -> "nome do TOM" -> "Tipo de folga (Ferias ou Compensa ...)"
    Dim Folga As New Scripting.Dictionary
    'Varre a manhã e tarde colocando o nome do tom na key e o tipo da folga no value
    Dim Row As Long
    For Row = start_vacation_line To end_vacation_line
        If (wsSrc.Cells(Row, morningColSrc) <> "" Or wsSrc.Cells(Row, afternoonColSrc) <> "") Then
            '              Nome TOM                     Nome da folga
            Folga(wsSrc.Cells(Row, "K").Value) = wsSrc.Cells(Row, "L").Value
        End If
    Next Row
    
    
    'Pega o dia da programação e as linhas de folga ----------------------
    Dim Folgas As String 'Pessoas que estarão de folga no dia
    Folgas = wsSrc.Cells(start_vacation_line, morningColSrc).Value & "; " & wsSrc.Cells(start_vacation_line, afternoonColSrc).Value
    'Pega o inicio do status dos técnicos
    Dim Status_Tech_Row As Long
    Status_Tech_Row = Final_Tarde + 3
    
    'Começa a preencher os dados dos técnicos
    Dim writeRow As Long
    writeRow = Status_Tech_Row
    
    Do While wsDst.Cells(writeRow, Ecol).Value <> ""
        
        Dim tech As String
        tech = wsDst.Cells(writeRow, Ecol).Value
        
        If Folga.Exists(tech) Then 'Caso esteja de folga
            wsDst.Cells(writeRow, Ccol).Value = Folga(tech)
        Else 'Caso esteja a trabalho insere a função para detectar plantão
        Dim FormulaLocal_str As String
        FormulaLocal_str = "=SE($F$2<=OCGR!$Q$4;" & _
                        "SE(E" & writeRow & "=" & wsSrc.Name & "!$J$" & end_vacation_line + 2 & ";""Sobreaviso (Laranja)"";" & _
                            "SE(E" & writeRow & "=" & wsSrc.Name & "!$J$" & end_vacation_line + 1 & ";""Sobreaviso (Vermelho)"";""A serviço""));" & _
                        "SE(E" & writeRow & "=" & wsSrc.Name & "!$K$" & end_vacation_line + 2 & ";""Sobreaviso (Laranja)"";" & _
                            "SE(E" & writeRow & "=" & wsSrc.Name & "!$K$" & end_vacation_line + 1 & ";""Sobreaviso (Vermelho)"";""A serviço"")))"
            wsDst.Cells(writeRow, Ccol).FormulaLocal = FormulaLocal_str
        End If
        
        writeRow = writeRow + 1
    Loop
    
End Sub

