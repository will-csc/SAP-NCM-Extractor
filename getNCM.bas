Sub getNCM()

    If Not IsObject(Application1) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set Application1 = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Application1.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    
    Set ws = ActiveSheet
    
    materialCont = ws.Range("A1000000").End(xlUp).Row
    
    For Each plant In Array("8802", "8801", "8800")
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "MM03"
        session.findById("wnd[0]").sendVKey 0
        For i = 2 To materialCont
            On Error Resume Next
            materialCode = ws.Cells(i, 1).Value
            session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = materialCode
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).Selected = True
            session.findById("wnd[1]/usr/chkUSRM1-AAUSW").SetFocus
            session.findById("wnd[1]/usr/chkUSRM1-AAUSW").Selected = True
            session.findById("wnd[2]/tbar[0]/btn[0]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = plant
            session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").caretPosition = 4
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            
            ncmValue = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2205/ctxtMARC-STEUC").Text
            If plant = "8801" Then
                ws.Cells(i, 2).Value = ncmValue
            ElseIf plant = "8802" Then
                ws.Cells(i, 3).Value = ncmValue
            Else
                ws.Cells(i, 4).Value = ncmValue
            End If
            
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            On Error GoTo 0
        Next i
        ncmValue = ""
        session.findById("wnd[0]/tbar[0]/btn[3]").press
    Next plant
    
End Sub
