VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private app As Word.Application
Private docs As Word.Documents
Private doc As Word.Document


Public Function get_docs() As Word.Documents
    Set get_docs = docs
End Function


Public Function get_doc() As Word.Document
    Set get_doc = doc
End Function

Private Sub Class_Initialize()


    
    If Not app Is Nothing Then
        Application.DisplayAlerts = False
        app.Quit False
        Application.DisplayAlerts = True
    End If
    
    
    
    
    Set app = New Word.Application
    Set docs = app.Documents
    app.Visible = True
End Sub

Private Sub Class_Terminate()


    If Not app Is Nothing Then
        Application.DisplayAlerts = False
        app.Quit False
        Application.DisplayAlerts = True
    End If
    

    Set app = Nothing
    Set docs = Nothing
    
    ' doc.Close False
End Sub


'Public Sub assign_new_word_doc(e As E_DOCS)'
'
'    Select Case e
'        Case BUCKET.DOC_DEL_CONF
'            ' kopiuj del conf worda
'            tworzenie_nowego_doca
'        Case BUCKET.DOC_GUIDE
'            tworzenie_nowego_doca
'        Case BUCKET.DOC_PART_ORDER_CONF
'            tworzenie_nowego_doca
'        Case Else
'            MsgBox "cos nie tak poszlo"
'    End Select
'End Sub

Public Sub connect_shape_with_new_doc(ByRef s As Excel.Shape, e As E_DOCS)



    ' Set doc = docs.Add
    ' app.Visible = True
    
    doc_name = ""
    
    Select Case e
        Case DOC_DEL_CONF
            doc_name = "DELIVERY CONFIRMATION"
        Case DOC_GUIDE
            doc_name = "Final GM Supplier EDI Requirements and EDI Onboarding Guide_2013"
        Case DOC_PART_ORDER_CONF
            doc_name = "PART ORDER CONFIRMATION"
        Case Else
            MsgBox "to nie moze sie pojawic w connect shape"
    End Select

    ' Range("U8").Select
    ' ActiveSheet.Shapes.Range(Array("Object 2")).Select
    ' Selection.Verb Verb:=xlOpen
    s.Parent.Activate
    s.Select
    Selection.Verb Verb:=xlOpen
    
    wait_for_opening Int(docs.Count)
    
    For x = 1 To docs.Count
        
        If CStr(docs(x).name) Like "*" & CStr(BUCKET.FILENAME_FOR_DOC) & "*" Then
            Set doc = docs(x)
            Application.DisplayAlerts = False
            doc.SaveAs2 ThisWorkbook.Path & Application.PathSeparator & doc_name & ".docx"
            
            doc.Close False
            Set doc = Nothing
            docs.Open ThisWorkbook.Path & Application.PathSeparator & doc_name & ".docx"
            wait_for_opening Int(docs.Count)
            
            Set doc = docs(CStr(doc_name) & ".docx")
            Application.DisplayAlerts = True
            Exit For
        End If
    Next x

End Sub


Public Sub wait_for_opening(i As Integer)
    Do
    Loop Until docs.Count = i
End Sub


Public Function getfullname() As String
    getfullname = CStr(doc.FullName)
End Function
