VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlookHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private o As Outlook.Application
Private aitem As Outlook.AppointmentItem
Private mailItem As Outlook.mailItem


Private str_to As String
Private str_body As String
Private str_title As String
Private str_script As String

Public table_in_html  As String
Private fs As Object


Private w As Workbook
Private sh As Worksheet
Private source_sh As Worksheet


Private wordh As WordHandler
Private my_shape As Excel.Shape

Private cfg As InitConfigHandler

Public Sub connect_with_init_cfg(mcfg As InitConfigHandler)
    Set cfg = mcfg
End Sub



Public Function getAttachedFiles() As Outlook.Attachments
    
    Set getAttachedFiles = mailItem.Attachments
End Function

Private Sub Class_Initialize()
    Set o = New Outlook.Application
    
    table_in_html = ""
End Sub


Public Sub set_str_body(s As String)
    str_body = CStr(s)
End Sub

Public Sub set_str_title(s As String)
    str_title = CStr(s)
End Sub

Public Sub set_str_script(s As String)
    str_script = CStr(s)
End Sub

Public Sub set_str_to(s As String)
    str_to = CStr(s)
End Sub



Public Sub save_html_item(ByVal oh_tmp As OutlookHandler)

    Set mailItem = o.CreateItem(olMailItem)

    With mailItem
        .Subject = str_title
        
        'Debug.Print "<html><head>" & _
        '    "<title>" & str_title & "</title>" & _
        '    "<script>" & str_script & "</script>" & _
        '    "</head>" & _
        '    "<body>" & str_body & "</body></html>"
        
        
        
        .HTMLBody = "<html><head>" & _
            "<title>" & str_title & "</title>" & _
            "<script>" & str_script & "</script>" & _
            "</head>" & _
            "<body>" & str_body & "</body></html>"
        .BodyFormat = olFormatHTML
        .To = CStr(BUCKET.TO_TXT)
        
        ' attaching files
        ' ==================================================
        If oh_tmp Is Nothing Then
        
        
        
        
            ' LABELKI
            ' ---------------------------------------------------------------------
            
            If cfg.att_labels Then
            
                Dim iter_sh As Worksheet
                For Each iter_sh In ThisWorkbook.Sheets
                
                    If iter_sh.name Like CStr(BUCKET.LABEL_SH_NAME) & "*" Then
                
                        Set w = Workbooks.Add
                        Set source_sh = iter_sh
                        source_sh.Copy Before:=w.Sheets(1)
                        
                        Application.DisplayAlerts = False
                        w.SaveAs ThisWorkbook.Path & Application.PathSeparator & CStr(iter_sh.name), 51, , , , , , xlLocalSessionChanges
                        Application.DisplayAlerts = True
                        
                        
                        .Attachments.Add w.FullName
                        w.Close False
                    End If
                Next iter_sh
            End If
            
            ' ---------------------------------------------------------------------
            
            
            ' BLUE TAG
            ' ---------------------------------------------------------------------
            
            If cfg.att_blue_tag Then
            
                ' Dim iter_sh As Worksheet
                Set iter_sh = ThisWorkbook.Sheets(BUCKET.BLUE_TAG_SH_NM)
        
                Set w = Workbooks.Add
                Set source_sh = iter_sh
                source_sh.Copy Before:=w.Sheets(1)
                
                Application.DisplayAlerts = False
                w.SaveAs ThisWorkbook.Path & Application.PathSeparator & CStr(iter_sh.name), 51, , , , , , xlLocalSessionChanges
                Application.DisplayAlerts = True
                
                
                .Attachments.Add w.FullName
                w.Close False
            End If
            
            ' ---------------------------------------------------------------------
            
            ' PPQP BIW LABEL
            ' ---------------------------------------------------------------------
            
            If cfg.att_ppqp_label Then
            
                ' Dim iter_sh As Worksheet
                Set iter_sh = ThisWorkbook.Sheets(BUCKET.PPQP_LABEL_SH_NM)
        
                Set w = Workbooks.Add
                Set source_sh = iter_sh
                source_sh.Copy Before:=w.Sheets(1)
                
                Application.DisplayAlerts = False
                w.SaveAs ThisWorkbook.Path & Application.PathSeparator & CStr(iter_sh.name), 51, , , , , , xlLocalSessionChanges
                Application.DisplayAlerts = True
                
                
                .Attachments.Add w.FullName
                w.Close False
            End If
            
            ' ---------------------------------------------------------------------
            
            
            ' excel - survey
            ' ---------------------------------------------------------------------
            
            If cfg.att_survey Then
            
                Set w = Workbooks.Add
                Set source_sh = ThisWorkbook.Sheets(BUCKET.SURVEY_SH_NAME)
                source_sh.Copy Before:=w.Sheets(1)
                
                Application.DisplayAlerts = False
                w.SaveAs ThisWorkbook.Path & Application.PathSeparator & CStr(BUCKET.SURVEY_SH_NAME), 51, , , , , , xlLocalSessionChanges
                Application.DisplayAlerts = True
                
                
                .Attachments.Add w.FullName
                w.Close False
                
            
            End If
            ' ---------------------------------------------------------------------
            
            
            
            
            ' word - del conf
            ' ---------------------------------------------------------------------
            
            If cfg.att_del_conf Then
            
                Set source_sh = ThisWorkbook.Sheets(BUCKET.DEL_CONF_SH_NAME)
                Set my_shape = source_sh.Shapes(1)
                
                Set wordh = New WordHandler
                wordh.connect_shape_with_new_doc my_shape, DOC_DEL_CONF
                
                .Attachments.Add wordh.getfullname()
                Set wordh = Nothing
            End If
            
            ' ---------------------------------------------------------------------
            
            ' word - ppqp action plan
            ' ---------------------------------------------------------------------
            
            If cfg.att_ppqp_action_plan Then
            
                Set source_sh = ThisWorkbook.Sheets(BUCKET.PPQP_ACTION_PLAN_SH_NM)
                Set my_shape = source_sh.Shapes(1)
                
                Set wordh = New WordHandler
                wordh.connect_shape_with_new_doc my_shape, DOC_PPQP_1_Quality_Warrant_AND_PPQP_2_Corrective_Action_Plan
                
                .Attachments.Add wordh.getfullname()
                Set wordh = Nothing
            End If
            
            ' ---------------------------------------------------------------------
            
            ' word - guide
            ' ---------------------------------------------------------------------
            If cfg.att_guide Then
                Set source_sh = ThisWorkbook.Sheets(BUCKET.GUIDE_SH_NAME)
                Set my_shape = source_sh.Shapes(1)
                
                Set wordh = New WordHandler
                wordh.connect_shape_with_new_doc my_shape, DOC_GUIDE
                
                .Attachments.Add wordh.getfullname()
                Set wordh = Nothing
            End If
            ' ---------------------------------------------------------------------
            
            
            ' word - part order conf
            ' ---------------------------------------------------------------------
                        
            If cfg.att_part_order_conf Then
                Set source_sh = ThisWorkbook.Sheets(BUCKET.PART_ORDER_CONF_SH_NAME)
                Set my_shape = source_sh.Shapes(1)
                
                Set wordh = New WordHandler
                wordh.connect_shape_with_new_doc my_shape, DOC_PART_ORDER_CONF
                
                .Attachments.Add wordh.getfullname()
                Set wordh = Nothing
                
            End If
            ' ---------------------------------------------------------------------
            
            
            
            ' ---------------------------------------------------------------------
            
            ' order - as excel file.
            ' ---------------------------------------------------------------------
            
            If cfg.att_material_order And cfg.radio_material_order_as_excel Then
            
                Set w = Workbooks.Add
                Set source_sh = ThisWorkbook.Sheets(BUCKET.MATERIAL_ORDER_SH_NAME)
                source_sh.Copy Before:=w.Sheets(1)
                
                Application.DisplayAlerts = False
                w.SaveAs ThisWorkbook.Path & Application.PathSeparator & CStr(source_sh.Range("TO")) & " _ORDER", 51, , , , , , xlLocalSessionChanges
                Application.DisplayAlerts = True
                
                
                .Attachments.Add w.FullName
                w.Close False
            End If
            ' ---------------------------------------------------------------------
            
            
            
            ' order as HTM file
            ' ---------------------------------------------------------------------
            If cfg.att_material_order And cfg.radio_material_order_as_html Then
                If table_in_html <> "" Then
                
                    path_with_filename = ThisWorkbook.Path & Application.PathSeparator & "orders.htm"
                    
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    Set a = fs.CreateTextFile(path_with_filename, True)
                    a.WriteLine CStr(table_in_html)
                    a.Close
                    
                    .Attachments.Add path_with_filename
                End If
            End If
            ' ---------------------------------------------------------------------
            
            
            
            
            ' ==================================================
        Else
            
            'NOK
            'Dim attitem As Attachment
            'For Each attitem In oh_tmp.getAttachedFiles()
            '
            '    .Attachments.Add ThisWorkbook.Path & Application.PathSeparator & CStr(attitem)
            'Next attitem
        End If
        
        .Save
        
        
        
        
        ' .Display
    End With

End Sub
