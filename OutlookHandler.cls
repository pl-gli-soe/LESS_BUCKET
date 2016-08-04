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


Private Sub Class_Initialize()
    Set o = New Outlook.Application
    
    
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
        .To = str_to & CStr(BUCKET.TO_TXT)
        
        
        ' attaching files
        ' ==================================================
        
        
        ' ==================================================
        
        
        .Save
        
        ' .Display
    End With

End Sub
