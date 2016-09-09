VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private supp_nm As String
Private duns As String

Private data_orderu As String

Private mrd As String
Private proj As String
Private phase As String
Private pickup As String
Private conf_req As String

Private pns_q As Dictionary
Private pns_nm As Dictionary

Public Function how_many_orders() As Long
    how_many_orders = CLng(pns_q.Count)
End Function


Public Function orders_to_string() As String
    
    tmp = "<table style=" & Chr(34) & "border: 2px solid black; width: 800px; border-collapse: collapse;" & Chr(34) & ">" & Chr(10)
    
    tmp = tmp & "<tr style=" & Chr(34) & "border: 1px solid black;" & Chr(34) & ">" & Chr(10)
        
        tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 4px;" & Chr(34) & ">" & "Pos" & "</th>" & Chr(10)
        tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 4px;" & Chr(34) & ">" & "Part Number" & "</th>" & Chr(10)
        tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 4px;" & Chr(34) & ">" & "Part Name" & "</th>" & Chr(10)
        tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 4px;" & Chr(34) & ">" & "Qty" & "</th>" & Chr(10)
    
    tmp = tmp & "</tr>"
    
    
    inc = 1
    
    For Each xkey In pns_q.Keys
    
        tmp = tmp & "<tr style=" & Chr(34) & "border: 1px solid black;" & Chr(34) & ">" & Chr(10)
        
            tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 2px;" & Chr(34) & ">" & CStr(inc) & "</th>" & Chr(10)
            tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 2px;" & Chr(34) & ">" & xkey & "</th>" & Chr(10)
            tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 2px;" & Chr(34) & ">" & pns_nm(xkey) & "</th>" & Chr(10)
            tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 2px;" & Chr(34) & ">" & pns_q(xkey) & "</th>" & Chr(10)
        
        tmp = tmp & "</tr>"
        
        inc = inc + 1
    Next xkey
    
    tmp = tmp & Chr(10) & "</table> <br/><br/>"
    
    
    orders_to_string = tmp
End Function

Public Sub set_name(suppnm As String, mduns As String)
    
    supp_nm = CStr(suppnm)
    duns = CStr(mduns)
End Sub

Public Sub set_info_from_details(d As Worksheet)
    'Private mrd As Date
    'Private proj As String
    'Private phase As String
    'Private pickup As String
    'Private conf_req As String
    
    'If IsDate(d.Range("mrd_date")) Then
    '    mrd = CStr(d.Range("mrd_date"))
    'Else
    '    mrd = ""
    'End If
    
    proj = CStr(d.Range("project"))
    phase = CStr(d.Range("faza"))
    pickup = CStr(d.Range("pickup_date"))
    mrd = CStr(d.Range("mrd"))
    data_orderu = CStr(Format(Date, "yyyy-mm-dd"))
    
    
    
    
End Sub


Public Sub put_orders_into_excel_worksheet(sh As Worksheet)


    sh.Range("B3:B7") = ""


    sh.Range("TO").Value = CStr(supp_nm)
    sh.Range("SF_DUNS").Value = CStr(duns)
    sh.Range("ORDER_DATE").Value = CStr(data_orderu)
    sh.Range("PROJ").Value = CStr(proj)
    sh.Range("PHASE").Value = CStr(phase)
    sh.Range("MRD").Value = CStr(mrd)
    sh.Range("PICKUP_DATE").Value = Left(CStr(Format(pickup, "yyyy-mm-dd")), 10)
    
    
    make_podpis_na_bazie_ostatniej_textboxa sh.Range("FROM")
    
    
    wiersz = 19
    
    
    ' super byle jak, ale bedzie dzialac na bank
    ' ---------------------------------------------
    sh.Range("A19:E10000").Clear
    ' ---------------------------------------------
    
    
    For Each xkey In pns_q.Keys
    
        sh.Cells(wiersz, 1) = wiersz - 18
        sh.Cells(wiersz, 2) = xkey
        sh.Cells(wiersz, 3) = pns_nm(xkey)
        sh.Cells(wiersz, 4) = pns_q(xkey)
    
        wiersz = wiersz + 1
    Next xkey
    
    If wiersz > 19 Then
        thinlines sh.Range("A19:D" & CStr(wiersz - 1))
    End If
End Sub


Private Sub thinlines(r As Range)
'
' thinlines Macro
'

'
    r.Borders(xlDiagonalDown).LineStyle = xlNone
    r.Borders(xlDiagonalUp).LineStyle = xlNone
    With r.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


Public Sub make_podpis_na_bazie_ostatniej_textboxa(ByRef r As Range)

    Dim hn As HtmlNode
    Set hn = New HtmlNode
    
    ' FTR
    'hn.add_node_with_embed_word "footer_s", "div", "footers", "font-size:14px;", _
    '    ThisWorkbook.Sheets("info").Shapes.Range(Array("FTR"))
    
    sig_from_bottom_embed = _
        hn.parse_to_raw_txt_from_embedobject(ThisWorkbook.Sheets("info").Shapes.Range(Array("FTR")))
    
    arr = Split(sig_from_bottom_embed, Chr(10))
    
    q = 0
    For x = LBound(arr) To UBound(arr)
    
        ' zabezpieczenie przed wejsciem textu na dalsza czesc formularza
        ' q < 5
        If q < 5 Then
            r.Offset(q, 0) = arr(x)
        Else
            Exit For
        End If
        
        q = q + 1
    Next x
    
End Sub


Public Function get_name() As String
    get_name = CStr(supp_nm) & " " & CStr(duns)
End Function


Public Function czy_jest_synchro_na_pn(pn As String) As Boolean
    czy_jest_synchro_na_pn = False
    
    If pns_q.Exists(CStr(pn)) And pns_nm.Exists(CStr(pn)) Then
        czy_jest_synchro_na_pn = True
    Else
        czy_jest_synchro_na_pn = False
    End If
End Function


Public Sub change_pn_order_qty(pn As String, qty As Long)
    
    If pns_q.Exists(CStr(pn)) Then
        
        pns_q(pn) = CLng(qty)
    Else
        MsgBox "nie ma czego zmieniac"
    End If
End Sub


Public Sub remove_pn_item(pn As String)

    If pns_q.Exists(CStr(pn)) Then
    
        pns_q.Remove CStr(pn)
        pns_nm.Remove CStr(pn)
    Else
        MsgBox "nie ma takiego do usuniecia"
    End If
End Sub

Public Sub try_to_add_new_pn(pn As String, pn_nm As String, qty As Long)

    If Not pns_q.Exists(CStr(pn)) Then
        pns_q.Add CStr(pn), CLng(qty)
        pns_nm.Add CStr(pn), CStr(pn_nm)
    Else
        MsgBox "juz taki numer zostal dodany!"
    End If

End Sub

Public Sub clear_dictionary()
    
    pns_q.RemoveAll
    pns_nm.RemoveAll
End Sub


Private Sub Class_Initialize()
    Set pns_q = New Dictionary
    Set pns_nm = New Dictionary
    
    mrd = ""
    proj = ""
    phase = ""
    pickup = ""
    conf_req = ""

End Sub

Private Sub Class_Terminate()
    Set pns_q = Nothing
    Set pns_nm = Nothing
End Sub
