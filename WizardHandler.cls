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
Private nh2 As HtmlNode


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


Private Function sprawdz_czy_wybrany_plik_to_faktycznie_wizard() As Boolean
    
    sprawdz_czy_wybrany_plik_to_faktycznie_wizard = False
    
    
    If CStr(wizard_w.Sheets(BUCKET.MASTER_SH_NAME).Cells(1, 1)) = "PN" Then
        If CStr(wizard_w.Sheets(BUCKET.MASTER_SH_NAME).Cells(1, 2)) = "Alternative PN" Then
            If CStr(wizard_w.Sheets(BUCKET.MASTER_SH_NAME).Cells(1, 3)) = "PN Name" Then
                If CStr(wizard_w.Sheets(BUCKET.MASTER_SH_NAME).Cells(1, 4)) = "GPDS PN Name" Then
                
                    ' mysle ze labelki z czeterech kolumn wystarcza
                    sprawdz_czy_wybrany_plik_to_faktycznie_wizard = True
                End If
            End If
        End If
    End If
    
End Function


Private Function sprawdz_czy_wizard_nie_jest_przypadkiem_pusty() As Boolean
    
    
    
    If Len(CStr(wizard_w.Sheets(BUCKET.MASTER_SH_NAME).Cells(2, 1))) > 0 Then
        sprawdz_czy_wizard_nie_jest_przypadkiem_pusty = False
    Else
        sprawdz_czy_wizard_nie_jest_przypadkiem_pusty = True
    End If
End Function



Public Sub uruchom_dla_danych_funkcjonalnosc_w_wybranym_wizardzie(e As E_MASTER_MANDATORY_COLUMNS, twoj_deck As String, dns As String)


    ' dns jako dodatkowy filter zatem jesli jest pusty to znaczy ze filtru nie ma
    
    For Each w In Workbooks
        If w.name = Me.nazwa_aktywnego_pliku Then
            Me.set_wizard_w w
        End If
    Next w
    
    
    If Not sprawdz_czy_wybrany_plik_to_faktycznie_wizard() Then
        Me.set_wizard_w Nothing
    End If
    
    
    
    
    
    If Not wizard_w Is Nothing Then
    
    
    
        If Not sprawdz_czy_wizard_nie_jest_przypadkiem_pusty() Then
    
        
            ' wizard jest podlaczony
            ' pobierz arkusz master
            Set master = wizard_w.Sheets(BUCKET.MASTER_SH_NAME)
            Set details = wizard_w.Sheets(BUCKET.DETAILS_SH_NAME)
            
            
            ' teraz lecimy z uzupelnieniem kolekcji orderow
            
            Dim r As Range
            Set r = master.Range("A2")
            Set order_handler = Nothing
            
            Do
            
                If r.Offset(0, BUCKET.duns - 1) Like "*" & CStr(dns) & "*" Then
                    If r.Offset(0, BUCKET.fup_code - 1) = twoj_deck Then
                
                        If orders.Exists(CStr(r.Offset(0, BUCKET.duns - 1))) Then
                        
                            Set order_handler = orders(CStr(r.Offset(0, BUCKET.duns - 1)))
                            order_handler.try_to_add_new_pn CStr(r.Offset(0, BUCKET.pn - 1)), CStr(r.Offset(0, BUCKET.PN_Name - 1)), CLng(r.Offset(0, e - 1))
                        Else
                        
                            Set order_handler = New OrderHandler
                            
                            With order_handler
                                .set_name CStr(r.Offset(0, BUCKET.Supplier_Name - 1)), CStr(r.Offset(0, BUCKET.duns - 1))
                                .set_info_from_details details
                                .try_to_add_new_pn CStr(r.Offset(0, BUCKET.pn - 1)), CStr(r.Offset(0, BUCKET.PN_Name - 1)), CLng(r.Offset(0, e - 1))
                            End With
                            orders.Add CStr(r.Offset(0, BUCKET.duns - 1)), order_handler
                        End If
                    End If
                End If
                
                Set r = r.Offset(1, 0)
            
            Loop Until Trim(r) = ""
            
            
            
            Set order_handler = Nothing
            For Each xkey In orders.Keys()
            
                Set order_handler = orders(xkey)
            
                
                    
                
                ThisWorkbook.Sheets("info").Range("A3").Offset(1, 1) = ""
                work_with ThisWorkbook.Sheets("info").Range("A3").Offset(1, 1), order_handler
                Dim tmp_html_txt_for_table As String
                tmp_html_txt_for_table = CStr(ThisWorkbook.Sheets("info").Range("A3").Offset(1, 1))
                
                
                Set nh = New HtmlNode
                nh.init "start", "div", "main", ""
                
                With nh
                
                
                    ' main heading
                    DoEvents
                    .add_node "fh", "div", "firsth", "font-size:20px;", CStr(BUCKET.FIRST_HEADING)
                
                
                    ' general info (tutaj tez bedzie sie miesici tabela zamowien)
                    ' ----------------------------------------------------------------------------------------------------------------
                    
                    DoEvents
                    .add_node "gen_info_1", "div", "gi1", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A1"))
                    '.add_node "gen_ops", "div", "go", "font-size:14px;", CStr(ThisWorkbook.Sheets("info").Range("A1").Offset(1, 1))
                    .add_node_with_embed_word "gen_info_1_s", "div", "gi1s", "font-size:14px;", _
                        ThisWorkbook.Sheets("info").Shapes.Range(Array("GEN_INFO_1"))
                        
                        
                    ' tutaj dodajemy tabele zamowien
                    ' =====================================
                    ' =====================================
                    DoEvents
                    If order_handler.how_many_orders() <= G_LIMIT_ORDEROW_JAKO_TABELA_BEZPOSREDNIO_W_TRESCI_MAILA Then
                        .add_node "tb", "div", "so", "font-size:14px;", CStr(ThisWorkbook.Sheets("info").Range("A3").Offset(1, 1))
                    Else
                        ' no operation
                    End If
                    ' =====================================
                    ' =====================================
                    
                    DoEvents
                    .add_node "gen_info_2", "div", "gi2", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A3"))
                    .add_node_with_embed_word "gen_info_2_s", "div", "gi2s", "font-size:14px;", _
                        ThisWorkbook.Sheets("info").Shapes.Range(Array("GEN_INFO_2"))
                    ' ----------------------------------------------------------------------------------------------------------------
                    
                    DoEvents
                    .add_node "footer", "div", "ftr", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A17"))
                    ' FTR
                    .add_node_with_embed_word "footer_s", "div", "footers", "font-size:14px;", _
                        ThisWorkbook.Sheets("info").Shapes.Range(Array("FTR"))
                    
                    

                    
                End With
                
                'Set nh2 = nh
                'With nh2
                    'DoEvents
                    '.add_node "ship_info_and_new_labels", "div", "sil", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A5"))
                    ' SHIP_INFO_NEW_LBLS
                    '.add_node_with_embed_word "ship_info_and_new_labels_s", "div", "sis", "font-size:14px;", _
                    '    ThisWorkbook.Sheets("info").Shapes.Range(Array("SHIP_INFO_NEW_LBLS"))
                    
                    
                    'DoEvents
                    '.add_node "q_rqms", "div", "qrqms", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A7"))
                    '' Q_RQMS
                    '.add_node_with_embed_word "q_rqms_s", "div", "qrqmss", "font-size:14px;", _
                    '    ThisWorkbook.Sheets("info").Shapes.Range(Array("Q_RQMS"))
                '
                '    DoEvents
                '    .add_node "pymnts", "div", "pymnts", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A9"))
                '    ' PYMNTS
                '    .add_node_with_embed_word "pymnts_s", "div", "pymntss", "font-size:14px;", _
                '        ThisWorkbook.Sheets("info").Shapes.Range(Array("PYMNTS"))
                '
                '    DoEvents
                '    .add_node "contact_details", "div", "condet", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A11"))
                '    ' CNT_DETAILS
                '    .add_node_with_embed_word "cd_s", "div", "cds", "font-size:14px;", _
                '        ThisWorkbook.Sheets("info").Shapes.Range(Array("CNT_DETAILS"))
                '
                '    DoEvents
                '    .add_node "edi_issues", "div", "edii", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A13"))
                '    ' EDI_ISSUES
                '    .add_node_with_embed_word "ei_s", "div", "ediis", "font-size:14px;", _
                '        ThisWorkbook.Sheets("info").Shapes.Range(Array("EDI_ISSUES"))'
                '
                '    DoEvents
                '    .add_node "survey", "div", "srvy", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A15"))
                '    ' SRVY
                '    .add_node_with_embed_word "survey_s", "div", "surveys", "font-size:14px;", _
                '        ThisWorkbook.Sheets("info").Shapes.Range(Array("SRVY"))
                '
                '    DoEvents
                '    .add_node "footer", "div", "ftr", "font-size:16px;", CStr(ThisWorkbook.Sheets("info").Range("A17"))
                '    ' FTR
                '    .add_node_with_embed_word "footer_s", "div", "footers", "font-size:14px;", _
                '        ThisWorkbook.Sheets("info").Shapes.Range(Array("FTR"))
                'End With
                
                
                Set oh = New OutlookHandler
                
                oh.table_in_html = CStr(tmp_html_txt_for_table)
                
                ' oh.set_str_body CStr(master.Range("A2"))
                'oh.set_str_title CStr(ThisWorkbook.Sheets("info").Range("B1")) & _
                '    " " & CStr(order_handler.get_name) & _
                '    '" " & CStr(details.Range("plant")) & _
                '    '" " & CStr(details.Range("project")) & _
                '    " " & CStr(details.Range("faza")) & _
                '    " " & CStr(details.Range("my"))
                
                
                oh.set_str_title CStr(ThisWorkbook.Sheets("info").Range("B1")) & ", " & _
                    " " & CStr(order_handler.get_name) & ", " & _
                    " " & CStr(details.Range("faza")) & _
                    " " & CStr(details.Range("my"))
                
                
                With oh
                    .set_str_to CStr(order_handler.get_name)
                    .set_str_script ""
                    
                    
                    ' cale text boxy wpadaja tutaj
                    ' wystarczy od wersji 0.94 jedno nh bez przerzucania danych
                    ' 2016-08-26
                    ' ============================================================
                    ' ============================================================
                    .set_str_body CStr(nh.to_string)
                    ' ============================================================
                    ' ============================================================
                End With
                
                oh.save_html_item
                
                
            Next xkey
            
            
            reassign_txt_on_info_sheet
    
            ThisWorkbook.Sheets(BUCKET.INFO_SH_NAME).Activate
            MsgBox "done!"
        Else
            MsgBox "wybrany wizard jest pusty!"
        End If
        
    Else
        MsgBox "ref dla arkusza wizarda nie zostal przechwycony w obiekcie klasy podmien handler - niewlasciwy wzorzec"
    End If
End Sub


Private Sub reassign_txt_on_info_sheet()

    ThisWorkbook.Sheets("info").Range("A1").Offset(1, 1).Value = BUCKET.DESC_TXT
    ThisWorkbook.Sheets("info").Range("A3").Offset(1, 1).Value = BUCKET.DESC_TXT
    ThisWorkbook.Sheets("info").Range("A5").Offset(1, 1).Value = BUCKET.DESC_TXT
    ThisWorkbook.Sheets("info").Range("A7").Offset(1, 1).Value = BUCKET.DESC_TXT
End Sub


Private Sub work_with(ByRef ra3 As Range, ByRef orderh As OrderHandler)


    ' HTML do jednej komorki
    With orderh
        ' ra3.Value = .get_name()
        ra3.Value = CStr(ra3.Value) & "<br/><br/>"
        ra3.Value = CStr(ra3.Value) & CStr(.orders_to_string())
    End With
    
    
    ' do arkusza material order
    With orderh
        .put_orders_into_excel_worksheet ThisWorkbook.Sheets(BUCKET.MATERIAL_ORDER_SH_NAME)
    End With
End Sub
