VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EsteLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    ' Verifica o nome da aba ativada
    Select Case Sh.Name
        
        Case "OCGR - DIÁRIA"
            Call Prog_Diária("OCGR", Sh.Name)
            Call techs_state("OCGR", Sh.Name)
            Application.StatusBar = "Atualizada planilha: " & Sh.Name
            
        Case "DIÁRIA - TLG"
            Call Prog_Diária("TLG", Sh.Name)
            Call techs_state("TLG", Sh.Name)
            Application.StatusBar = "Atualizada planilha: " & Sh.Name
            
        Case "DIÁRIA - MCGR"
            Call Prog_Diária("MCGR", Sh.Name)
            Call techs_state("MCGR", Sh.Name)
            Application.StatusBar = "Atualizada planilha: " & Sh.Name
            
        Case "SEMANAL"
            Call mdl_Prog_Semanal.Updated_data
            Call fill_work("OCGR")
            Call fill_work("MCGR")
            Call fill_work("TLG")
        ' Case Else
            ' Aqui você pode colocar código para as abas que NÃO são as de cima
            ' (Opcional: se não colocar nada, o código apenas ignora outras abas)
            
    End Select

End Sub
