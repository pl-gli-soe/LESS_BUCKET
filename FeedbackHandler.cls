VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeedbackHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private oh As OutlookHandler

Private ile_orderow_zgromadzilem As Long
Private ile_orderow_z_pustym_dunsem As Long

Private str_pnow_ktore_sa_z_pustego_dunsu As String

Private ile_pnow_z_wizarda As Long
Private selected_ile_dunsow_z_wizarda As Long

Private duns_dictionary As Dictionary


Private n As HtmlNode



Public Property Get p_oh() As OutlookHandler
    Set get_oh = oh
End Property

Public Property Let p_oh(moh As OutlookHandler)
    Set oh = moh
End Property

Public Sub collect_info_from_wizard_master_sh(m As Worksheet)

    ile_pnow_z_wizarda = 0
    selected_ile_dunsow_z_wizarda = 0
        
    Dim r As Range
    Set r = m.Cells(2, BUCKET.pn)
    
    Do
        If Not r.EntireRow.Hidden Then
            If Trim(m.Cells(r.Row, BUCKET.duns)) <> "" Then
                If duns_dictionary.Exists(Trim(m.Cells(r.Row, BUCKET.duns))) Then
                Else
                    duns_dictionary.Add Trim(m.Cells(r.Row, BUCKET.duns)), 0
                End If
            End If
            ile_pnow_z_wizarda = ile_pnow_z_wizarda + 1
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    
    selected_ile_dunsow_z_wizarda = duns_dictionary.Count
End Sub

Public Sub collect_info_about_proxy_orders_collection_of_objects(ByRef orders As Dictionary)

    ' przelecimy teraz przez wszystkie zgromadzone ordery i podliczymy ilosci
    ile_orderow_zgromadzilem = orders.Count
    
    ile_orderow_z_pustym_dunsem = podlicz_puste_dunse(orders)
    
End Sub

Private Function podlicz_puste_dunse(ByRef orders As Dictionary) As Long
    Dim tmp As Long
    tmp = 0
    
    Dim i As OrderHandler
    For Each k In orders.Keys
    
        If Trim(k) = "" Then
            Set i = orders(k)
            str_pnow_ktore_sa_z_pustego_dunsu = str_pnow_ktore_sa_z_pustego_dunsu _
                & Chr(10) & Chr(10) & _
                i.orders_to_string()
                
            tmp = tmp + 1
        End If
    Next
    
    podlicz_puste_dunse = tmp
End Function


Public Sub generate_mail_report_at_the_end(ByRef orders As Dictionary)
    
    oh.set_str_title "BUCKET REPORT " & CStr(Now)
    oh.set_str_to "REPORT"
    Dim tmp_body As String
    Set n = New HtmlNode
    
    strtmpduns = "ILOSC DUNSOW W WIZARDZIE Z WYBRANEGO WYFILTROWANEGO PRZEDZIALU: " & CStr(selected_ile_dunsow_z_wizarda)
    strtmppn = "ILOSC PNOW W WIZARDZIE Z WYBRANEGO WYFILTROWANEGO PRZEDZIALU: " & CStr(ile_pnow_z_wizarda)
    
    strtmporderduns = "ILOSC MAILI (ORDEROW), WARTOSC TA POWINNA BYC ROWNA ILOSCI DUNSOW W WIZARDZIE: " & CStr(ile_orderow_zgromadzilem)
    strtmporderemptyduns = "ILOSC TZW. PUSTYCH DUNSOW (ile NOKow): " & CStr(ile_orderow_z_pustym_dunsem)
    
    n.add_node "r1", "div", "r1", "font-size:14px;", BUCKET.FIRST_LINE_IN_BUCKET_REPORT
    
    n.add_node "r2", "div", "r2", "font-size:12px;", "Stats: "
    n.add_node "r3", "div", "r3", "font-size:12px;", CStr(strtmpduns)
    n.add_node "r4", "div", "r4", "font-size:12px;", CStr(strtmppn)
    n.add_node "r5", "div", "r5", "font-size:12px;", CStr(strtmporderduns)
    n.add_node "r6", "div", "r6", "font-size:12px;", CStr(strtmporderemptyduns)
    
    n.add_node "ri7", "div", "ri7", "font-size: 10px;", CStr("DUNSes from ORDERS")
    n.add_node "r7", "div", "r7", "font-size:8px;", CStr(n.get_table_from_dictionary(orders))
    
    n.add_node "ri8", "div", "ri8", "font-size: 10px;", CStr("DUNSes from Wizard")
    n.add_node "r8", "div", "r8", "font-size:8px;", CStr(n.get_table_from_dictionary(duns_dictionary))
    
    tmp_body = n.to_string
    
    
    
    
    tmp_body = tmp_body & _
        "<br/>Tabele osieroconych PN (to sa NOKi, ktore musisz zweryfikowac, gdyz nie posiadaja prawidlowego orderu): <br/>" & _
        str_pnow_ktore_sa_z_pustego_dunsu & "<br/>"
        
        
    oh.set_str_body tmp_body
    oh.save_html_item Nothing
    
End Sub

Private Sub Class_Initialize()
    Set oh = New OutlookHandler
    
    ile_orderow_zgromadzilem = 0
    ile_orderow_z_pustym_dunsem = 0
    
    Set duns_dictionary = New Dictionary
    
    
    str_pnow_ktore_sa_z_pustego_dunsu = ""
End Sub

Private Sub Class_Terminate()
    Set oh = Nothing
    Set duns_dictionary = Nothing
End Sub

