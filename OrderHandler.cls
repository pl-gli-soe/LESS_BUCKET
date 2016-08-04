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

Private pns_q As Dictionary
Private pns_nm As Dictionary




Public Function orders_to_string() As String
    
    tmp = "<table style=" & Chr(34) & "border: 2px solid black; width: 800px; border-collapse: collapse;" & Chr(34) & ">" & Chr(10)
    
    tmp = tmp & "<tr style=" & Chr(34) & "border: 1px solid black;" & Chr(34) & ">" & Chr(10)
        
        tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 4px;" & Chr(34) & ">" & "Part Number" & "</th>" & Chr(10)
        tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 4px;" & Chr(34) & ">" & "Part Name" & "</th>" & Chr(10)
        tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 4px;" & Chr(34) & ">" & "Qty" & "</th>" & Chr(10)
    
    tmp = tmp & "</tr>"
    
    
    
    
    For Each xkey In pns_q.Keys
    
        tmp = tmp & "<tr style=" & Chr(34) & "border: 1px solid black;" & Chr(34) & ">" & Chr(10)
        
            tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 2px;" & Chr(34) & ">" & xkey & "</th>" & Chr(10)
            tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 2px;" & Chr(34) & ">" & pns_nm(xkey) & "</th>" & Chr(10)
            tmp = tmp & "<th style=" & Chr(34) & "border: 1px solid black; padding: 2px;" & Chr(34) & ">" & pns_q(xkey) & "</th>" & Chr(10)
        
        tmp = tmp & "</tr>"
    Next xkey
    
    tmp = tmp & Chr(10) & "</table> <br/><br/>"
    
    
    orders_to_string = tmp
End Function

Public Sub set_name(suppnm As String, mduns As String)
    
    supp_nm = CStr(suppnm)
    duns = CStr(mduns)
End Sub


Public Function get_name() As String
    get_name = "DUNS #: " & CStr(duns) & ", Supplier Name: " & CStr(supp_nm)
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
End Sub

Private Sub Class_Terminate()
    Set pns_q = Nothing
    Set pns_nm = Nothing
End Sub
