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
Private aItem As Outlook.AppointmentItem
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



Public Sub save_html_item()

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
        
        
        ' LABELKA
        ' ---------------------------------------------------------------------
        
        Set w = Workbooks.Add
        Set source_sh = ThisWorkbook.Sheets(BUCKET.LABEL_SH_NAME)
        source_sh.Copy Before:=w.Sheets(1)
        
        Application.DisplayAlerts = False
        w.SaveAs ThisWorkbook.Path & Application.PathSeparator & CStr(BUCKET.LABEL_SH_NAME), 51, , , , , , xlLocalSessionChanges
        Application.DisplayAlerts = True
        
        
        .Attachments.Add w.FullName
        w.Close False
        
        ' ---------------------------------------------------------------------
        
        ' excel - survey
        ' ---------------------------------------------------------------------
        
        Set w = Workbooks.Add
        Set source_sh = ThisWorkbook.Sheets(BUCKET.SURVEY_SH_NAME)
        source_sh.Copy Before:=w.Sheets(1)
        
        Application.DisplayAlerts = False
        w.SaveAs ThisWorkbook.Path & Application.PathSeparator & CStr(BUCKET.SURVEY_SH_NAME), 51, , , , , , xlLocalSessionChanges
        Application.DisplayAlerts = True
        
        
        .Attachments.Add w.FullName
        w.Close False
        
        ' ---------------------------------------------------------------------
        
        
        
        
        ' word - del conf
        ' ---------------------------------------------------------------------
        
        Set source_sh = ThisWorkbook.Sheets(BUCKET.DEL_CONF_SH_NAME)
        Set my_shape = source_sh.Shapes(1)
        
        Set wordh = New WordHandler
        wordh.connect_shape_with_new_doc my_shape, DOC_DEL_CONF
        
        .Attachments.Add wordh.getfullname()
        Set wordh = Nothing
        ' ---------------------------------------------------------------------
        
        
        
        ' word - guide
        ' ---------------------------------------------------------------------
        
        Set source_sh = ThisWorkbook.Sheets(BUCKET.GUIDE_SH_NAME)
        Set my_shape = source_sh.Shapes(1)
        
        Set wordh = New WordHandler
        wordh.connect_shape_with_new_doc my_shape, DOC_GUIDE
        
        .Attachments.Add wordh.getfullname()
        Set wordh = Nothing
        ' ---------------------------------------------------------------------
        
        
        ' word - part order conf
        ' ---------------------------------------------------------------------
        
        Set source_sh = ThisWorkbook.Sheets(BUCKET.PART_ORDER_CONF_SH_NAME)
        Set my_shape = source_sh.Shapes(1)
        
        Set wordh = New WordHandler
        wordh.connect_shape_with_new_doc my_shape, DOC_PART_ORDER_CONF
        
        .Attachments.Add wordh.getfullname()
        Set wordh = Nothing
        ' ---------------------------------------------------------------------
        
        
        
        ' ---------------------------------------------------------------------
        
        ' order - as excel file.
        ' ---------------------------------------------------------------------
        Set w = Workbooks.Add
        Set source_sh = ThisWorkbook.Sheets(BUCKET.MATERIAL_ORDER_SH_NAME)
        source_sh.Copy Before:=w.Sheets(1)
        
        Application.DisplayAlerts = False
        w.SaveAs ThisWorkbook.Path & Application.PathSeparator & CStr(BUCKET.MATERIAL_ORDER_SH_NAME), 51, , , , , , xlLocalSessionChanges
        Application.DisplayAlerts = True
        
        
        .Attachments.Add w.FullName
        w.Close False
        ' ---------------------------------------------------------------------
        
        
        
        ' order as HTM file
        ' ---------------------------------------------------------------------
        ' If table_in_html <> "" Then
        '
        '     path_with_filename = ThisWorkbook.Path & Application.PathSeparator & "orders.htm"
        '
        '     Set fs = CreateObject("Scripting.FileSystemObject")
        '     Set a = fs.CreateTextFile(path_with_filename, True)
        '     a.WriteLine CStr(table_in_html)
        '     a.Close
        '
        '     .Attachments.Add path_with_filename
        ' End If
        ' ---------------------------------------------------------------------
        
        
        
        
        ' ==================================================
        
        
        .Save
        
        
        
        
        ' .Display
    End With

End Sub
