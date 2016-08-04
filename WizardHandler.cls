VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WizardHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public nazwa_aktywnego_pliku As String
Private w As Workbook
Private wizard_w As Workbook


Private master As Worksheet
Private details As Worksheet


Private type_of_action As E_ACTION_TYPE


Private oh As OutlookHandler
Private nh As HtmlNode


Private orders As Dictionary
Private order_handler As OrderHandler


Public Sub setType(e As E_ACTION_TYPE)
    type_of_action = e
End Sub

Public Function getType() As E_ACTION_TYPE
    getType = type_of_action
End Function

Public Sub set_wizard_w(w As Workbook)
    Set wizard_w = w
End Sub


Private Sub Class_Initialize()
    nazwa_aktywnego_pliku = ""
    Set wizard_w = Nothing
    
    Set oh = New OutlookHandler
    Set orders = New Dictionary
End Sub



Public Sub uruchom_dla_danych_funkcjonalnosc_w_wybranym_wizardzie(e As E_MASTER_MANDATORY_COLUMNS, twoj_deck As String)

    
    For Each w In Workbooks
        If w.name = Me.nazwa_aktywnego_pliku Then
            Me.set_wizard_w w
        End If
    Next w
    
    
    If Not wizard_w Is Nothing Then
    
    
        ' wizard jest podlaczony
        ' pobierz arkusz master
        Set master = wizard_w.Sheets(o.MASTER_SH_NAME)
        Set details = wizard_w.Sheets(o.DETAILS_SH_NAME)
        
        
        ' teraz lecimy z uzupelnieniem kolekcji orderow
        
        Dim r As Range
        Set r = master.Range("A2")
        Set order_handler = Nothing
        
        Do
        
        
            If r.Offset(0, o.fup_code - 1) = twoj_deck Then
        
                If orders.Exists(CStr(r.Offset(0, o.duns - 1))) Then
                
                
                    
                
                    Set order_handler = orders(CStr(r.Offset(0, o.duns - 1)))
                    
                    
                    order_handler.try_to_add_new_pn CStr(r.Offset(0, o.pn - 1)), CStr(r.Offset(0, o.PN_Name - 1)), CLng(r.Offset(0, e - 1))
                Else
                    Set order_handler = New OrderHandler
                    With order_handler
                        .set_name CStr(r.Offset(0, o.Supplier_Name - 1)), CStr(r.Offset(0, o.duns - 1))
                        .try_to_add_new_pn CStr(r.Offset(0, o.pn - 1)), CStr(r.Offset(0, o.PN_Name - 1)), CLng(r.Offset(0, e - 1))
                    End With
                    orders.Add CStr(r.Offset(0, o.duns - 1)), order_handler
                End If
            End If
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        
        
        Set order_handler = Nothing
        For Each xkey In orders.Keys()
        
            Set order_handler = orders(xkey)
        
            ' oh.set_str_body CStr(master.Range("A2"))
            oh.set_str_title CStr(order_handler.get_name) & _
                " " & CStr(details.Parent.name) & _
                " " & CStr(details.Range("plant")) & _
                " " & CStr(details.Range("project")) & " " & _
                CStr(details.Range("faza"))
                
            
            ThisWorkbook.Sheets("info").Range("A3").Offset(1, 1) = ""
            work_with ThisWorkbook.Sheets("info").Range("A3").Offset(1, 1), order_handler
            
            
            Set nh = New HtmlNode
            nh.init "start", "div", "main", ""
            
            With nh
                .add_node "gen_info", "div", "gi", "font-size:20px;", CStr(ThisWorkbook.Sheets("info").Range("A1"))
                .add_node "gen_ops", "div", "go", "font-size:14px;", CStr(ThisWorkbook.Sheets("info").Range("A1").Offset(1, 1))
                
                
                .add_node "ship_info", "div", "si", "font-size:20px;", CStr(ThisWorkbook.Sheets("info").Range("A3"))
                .add_node "ship_ops", "div", "so", "font-size:14px;", CStr(ThisWorkbook.Sheets("info").Range("A3").Offset(1, 1))
                
                .add_node "lbls", "div", "lbls", "font-size:20px;", CStr(ThisWorkbook.Sheets("info").Range("A5"))
                .add_node "lblso", "div", "lblso", "font-size:14px;", CStr(ThisWorkbook.Sheets("info").Range("A5").Offset(1, 1))
                
                .add_node "pymtns", "div", "pymtns", "font-size:20px;", CStr(ThisWorkbook.Sheets("info").Range("A7"))
                .add_node "pymtnso", "div", "pymtnso", "font-size:14px;", CStr(ThisWorkbook.Sheets("info").Range("A7").Offset(1, 1))
            End With
            
            With oh
                .set_str_script ""
                .set_str_body CStr(nh.to_string)
            End With
            
            oh.save_html_item
            
            
        Next xkey
        
        
        reassign_txt_on_info_sheet

        
    Else
        MsgBox "ref dla arkusza wizarda nie zostal przechwycony w obiekcie klasy podmien handler"
    End If
End Sub


Private Sub reassign_txt_on_info_sheet()

    ThisWorkbook.Sheets("info").Range("A1").Offset(1, 1).Value = o.DESC_TXT
    ThisWorkbook.Sheets("info").Range("A3").Offset(1, 1).Value = o.DESC_TXT
    ThisWorkbook.Sheets("info").Range("A5").Offset(1, 1).Value = o.DESC_TXT
    ThisWorkbook.Sheets("info").Range("A7").Offset(1, 1).Value = o.DESC_TXT
End Sub


Private Sub work_with(ByRef ra3 As Range, ByRef orderh As OrderHandler)

    With orderh
        ra3.Value = .get_name()
        ra3.Value = CStr(ra3.Value) & "<br/><br/>"
        ra3.Value = CStr(ra3.Value) & CStr(.orders_to_string())
    End With
End Sub
