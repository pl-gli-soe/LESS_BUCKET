VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HtmlNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private inner_nodes As Collection

Private n As HtmlNode

Private name As String
Private znacznik As String

Private str_body As String

Private pre As String
Private post As String


Private uniq_id As String


Public Function get_uniq_id() As String
    get_uniq_id = uniq_id
End Function


Public Sub init_with_embed_word(nm As String, m As String, Optional ajdi As String, Optional stajl As String, Optional body As Object)

    name = nm
    If Len(m) > 0 Then
    
    
        ' tutaj bardzo wazne styl nie moze zostac uruchomiony bez przypisania id
        If Len(ajdi) > 0 Then
            pre = "<" & CStr(m) & _
                " id=" & Chr(34) & CStr(ajdi) & Chr(34) & _
                " style=" & Chr(34) & CStr(stajl) & Chr(34) & " > "
        Else
            pre = "<" & CStr(m) & ">"
        End If
        post = "</" & CStr(m) & ">"
    Else
        pre = ""
        post = ""
    End If
    
    
    tmp = CStr(parse_to_txt_from_embedobject(body))
    If Len(tmp) > 0 Then
        str_body = CStr(tmp)
    End If
End Sub

Public Sub init_with_embed_img(nm As String, m As String, Optional ajdi As String, Optional stajl As String, Optional body As Object)
    name = nm
    If Len(m) > 0 Then
    
    
        ' tutaj bardzo wazne styl nie moze zostac uruchomiony bez przypisania id
        If Len(ajdi) > 0 Then
            pre = "<" & CStr(m) & _
                " id=" & Chr(34) & CStr(ajdi) & Chr(34) & _
                " style=" & Chr(34) & CStr(stajl) & Chr(34) & " > "
        Else
            pre = "<" & CStr(m) & ">"
        End If
        post = "</" & CStr(m) & ">"
    Else
        pre = ""
        post = ""
    End If
    
    
    tmp = CStr(img_embedobject(body))
    If Len(tmp) > 0 Then
        str_body = CStr(tmp)
    End If
End Sub


Public Function img_embedobject(Optional b As Object) As String
    img_embedobject = ""
    
    If Not IsMissing(b) Then
        
        ' sekcja sciagania textu z obiektu embed
        ' ======================================
        ''
        '
        
        ' ok teraz by trzeba bylo znalezc mozliwosc wyciagniecia textu z tego obiektu
        ' -----------------------------------------------------------------------------
        Dim ss As Excel.Shapes
        Dim sr As Excel.ShapeRange
        Dim s As Excel.Shape
        
        ' Debug.Print TypeName(b) ' ShapeRange
        Set sr = b
        ' Debug.Print sr.AutoShapeType ' -2
        If sr.Type = msoPicture Then
            ' Debug.Print "this is mso embed ole object" ' true
            ' ==================================================
            ''
            '
            
            Set s = sr.Item(1)
            ' a mogloby byc tak pieknie gdyby to ponizsze dzialalo od razu....
            ' s.SaveAsPicture "img.jpg"
            ' s.Select
            '
            'Selection.Export "i.jpg"
            ' Dim p As Picture
            
            
            
            
            
            img_embedobject = ""
            
            
        End If
    End If
End Function

Public Function parse_to_txt_from_embedobject(Optional b As Object) As String
    parse_to_txt_from_embedobject = ""
    
    If Not IsMissing(b) Then
        
        
        ' sekcja sciagania textu z obiektu embed
        ' ======================================
        ''
        '
        
        ' ok teraz by trzeba bylo znalezc mozliwosc wyciagniecia textu z tego obiektu
        ' -----------------------------------------------------------------------------
        Dim sr As Excel.ShapeRange
        Dim s As Excel.Shape
        
        ' Debug.Print TypeName(b) ' ShapeRange
        Set sr = b
        ' Debug.Print sr.AutoShapeType ' -2
        If sr.Type = msoEmbeddedOLEObject Then
            ' Debug.Print "this is mso embed ole object" ' true
            ' ==================================================
            ''
            '
            
            Set s = sr.Item(1)
            'Debug.Print s.name
            'Debug.Print TypeName(s.OLEFormat.Object)
            
            s.OLEFormat.Verb xlOpen
            
            Dim doc As Word.Document
            Dim docs As Word.Documents
            Dim app As Word.Application
            Dim data_object As MSForms.DataObject
            
        
            Set app = Nothing
            Do
                Set app = GetObject(, "Word.Application")
            Loop While app Is Nothing
            
            
            ' sekcja wyciagniecia tekstu z pliku
            ' =======================================
            
            ' przypominam ze txt trzeba wrzucic:
            ' parse_to_txt_from_embedobject
            ' ----------------------------------------------
            ''
            '
            Set docs = app.Documents
            Set doc = docs(1)
            
            
            
            Dim rr As Range
            
            
            Dim p As Word.Paragraph
            ' MsgBox .Paragraphs(1).Range.Text
            With doc
                For Each p In .Paragraphs
                    ' MsgBox p.Range.text
                    
                    
                    
                    
                    'Set rr = ThisWorkbook.Sheets(BUCKET.REGISTER_SH_NM).Range("A1")
                    'rr.PasteSpecial xlPasteAll
                    
                    tmp_txt = tmp_txt & "<br/>" & CStr(fnConvert2HTML(p.Range))
                    
                    
                Next p
                
            End With
            '
            ''
            ' ----------------------------------------------
            
            ' =======================================
            
            doc.Close False
            
            If Not app Is Nothing Then
                Application.DisplayAlerts = False
                app.Quit False
                Application.DisplayAlerts = True
            End If
            
            
            parse_to_txt_from_embedobject = tmp_txt
            
            
            
            ' teoretycznie udalo mi sie zlinkowac obiekt typu word
            ' teraz sekcja pracy na tekscie i odopowiednie przekopiowanie go do formatu html,
            ' ktory bedzie uzyteczny z perspektywy maila (kolory, rozmiar czcionki itd)
            ' -----------------------------------------------------------------------------------
            
            
            ' -----------------------------------------------------------------------------------
            
            '
            ''
            ' ==================================================
        End If
        
        
        ' -----------------------------------------------------------------------------
        
        '
        ''
        ' ======================================
    End If
    
End Function


Public Function parse_to_raw_txt_from_embedobject(Optional b As Object) As String
    parse_to_raw_txt_from_embedobject = ""
    
    If Not IsMissing(b) Then
        
        
        ' sekcja sciagania textu z obiektu embed
        ' ======================================
        ''
        '
        
        ' ok teraz by trzeba bylo znalezc mozliwosc wyciagniecia textu z tego obiektu
        ' -----------------------------------------------------------------------------
        Dim sr As Excel.ShapeRange
        Dim s As Excel.Shape
        
        ' Debug.Print TypeName(b) ' ShapeRange
        Set sr = b
        ' Debug.Print sr.AutoShapeType ' -2
        If sr.Type = msoEmbeddedOLEObject Then
            ' Debug.Print "this is mso embed ole object" ' true
            ' ==================================================
            ''
            '
            
            Set s = sr.Item(1)
            'Debug.Print s.name
            'Debug.Print TypeName(s.OLEFormat.Object)
            
            s.OLEFormat.Verb xlOpen
            
            Dim doc As Word.Document
            Dim docs As Word.Documents
            Dim app As Word.Application
            Dim data_object As MSForms.DataObject
            
        
            Set app = Nothing
            Do
                Set app = GetObject(, "Word.Application")
            Loop While app Is Nothing
            
            
            ' sekcja wyciagniecia tekstu z pliku
            ' =======================================
            
            ' przypominam ze txt trzeba wrzucic:
            ' parse_to_txt_from_embedobject
            ' ----------------------------------------------
            ''
            '
            Set docs = app.Documents
            Set doc = docs(1)
            
            
            Dim p As Word.Paragraph
            tmp_txt = ""
            With doc
                For Each p In .Paragraphs
                    tmp_txt = tmp_txt & p.Range.Text & Chr(10)
                Next p
                
                tmp_txt = tmp_txt & Chr(10)
            End With
            
            'Dim rr As Range
           '
            '
            'Dim p As Word.Paragraph
            '' MsgBox .Paragraphs(1).Range.Text
            'With doc
            '    For Each p In .Paragraphs
            '        ' MsgBox p.Range.text
            '
            '
            '
             '
             '       'Set rr = ThisWorkbook.Sheets(BUCKET.REGISTER_SH_NM).Range("A1")
             '       'rr.PasteSpecial xlPasteAll
             '
             '       tmp_txt = tmp_txt & "<br/>" & CStr(fnConvert2HTML(p.Range))
             ''
             '
             '   Next p
            '
            'End With
            '
            ''
            ' ----------------------------------------------
            
            ' =======================================
            
            doc.Close False
            
            If Not app Is Nothing Then
                Application.DisplayAlerts = False
                app.Quit False
                Application.DisplayAlerts = True
            End If
            
            
            parse_to_raw_txt_from_embedobject = tmp_txt
            
            
            
            ' teoretycznie udalo mi sie zlinkowac obiekt typu word
            ' teraz sekcja pracy na tekscie i odopowiednie przekopiowanie go do formatu html,
            ' ktory bedzie uzyteczny z perspektywy maila (kolory, rozmiar czcionki itd)
            ' -----------------------------------------------------------------------------------
            
            
            ' -----------------------------------------------------------------------------------
            
            '
            ''
            ' ==================================================
        End If
        
        
        ' -----------------------------------------------------------------------------
        
        '
        ''
        ' ======================================
    End If
    
End Function



Public Sub init(nm As String, m As String, Optional ajdi As String, Optional stajl As String, Optional body As String)


    uniq_id = ajdi

    name = nm
    If Len(m) > 0 Then
    
    
        ' tutaj bardzo wazne styl nie moze zostac uruchomiony bez przypisania id
        If Len(ajdi) > 0 Then
            pre = "<" & CStr(m) & _
                " id=" & Chr(34) & CStr(ajdi) & Chr(34) & _
                " style=" & Chr(34) & CStr(stajl) & Chr(34) & " > "
        Else
            pre = "<" & CStr(m) & ">"
        End If
        post = "</" & CStr(m) & ">"
    Else
        pre = ""
        post = ""
    End If
    
    If Len(body) > 0 Then
        str_body = body
    End If
End Sub


Public Function get_collection_of_inner_nodes() As Collection
    If inner_nodes.Count > 0 Then
        Set get_collection_of_inner_nodes = inner_nodes
    Else
        Set get_collection_of_inner_nodes = Nothing
    End If
End Function


Public Sub add_node(nm As String, m As String, Optional ajdi As String, Optional stajl As String, Optional body As String)
    Set n = New HtmlNode
    n.init nm, m, ajdi, stajl, body
    inner_add_node n
End Sub


Public Sub change_node(nm As String, m As String, Optional ajdi As String, Optional stajl As String, Optional body As String)
    Set n = New HtmlNode
    n.init nm, m, ajdi, stajl, body
    inner_change_node n
End Sub

Public Sub add_node_with_embed_word(nm As String, m As String, Optional ajdi As String, Optional stajl As String, Optional body As Object)
    Set n = New HtmlNode
    n.init_with_embed_word nm, m, ajdi, stajl, body
    inner_nodes.Add n
End Sub

Public Sub add_node_with_embed_img(nm As String, m As String, Optional ajdi As String, Optional stajl As String, Optional body As Object)
    Set n = New HtmlNode
    n.init_with_embed_img nm, m, ajdi, stajl, body
    inner_nodes.Add n
End Sub

Private Sub inner_change_node(n As HtmlNode)

    Dim inode As HtmlNode
    Dim tmp_coll As Collection
    Set tmp_coll = inner_nodes
    
    
    Set inner_nodes = Nothing
    Set inner_nodes = New Collection
    
    For Each inode In tmp_coll
    
        If CStr(inode.get_uniq_id) = CStr(n.get_uniq_id) Then
            inner_nodes.Add n
        Else
            inner_nodes.Add inode
        End If
    Next inode
    
    
End Sub

Private Sub inner_add_node(n As HtmlNode)

    inner_nodes.Add n
End Sub

Private Sub Class_Initialize()
    Set inner_nodes = Nothing
    Set inner_nodes = New Collection
End Sub

Private Sub Class_Terminate()
    Set inner_nodes = Nothing
End Sub




Public Function to_string() As String
    
    to_string = ""
    
    tmp = CStr(pre) & _
        str_body & _
        get_string_from_collection() & CStr(post) & "<br/>"
    
    to_string = CStr(tmp)
End Function

Public Function special_to_string_for_material_order_sheet() As String
    
    special_to_string_for_material_order_sheet = special_get_string_from_collection()
End Function


Private Function get_string_from_collection() As String
    get_string_from_collection = ""
    
    On Error Resume Next
    If inner_nodes.Count > 0 Then
        Dim nodeItem As HtmlNode
        For Each nodeItem In inner_nodes
            
            get_string_from_collection = get_string_from_collection & nodeItem.to_string & "<br/>"
        Next nodeItem
    End If
End Function

Private Function special_get_string_from_collection() As String
    
    special_get_string_from_collection = ""
    
    On Error Resume Next
    If inner_nodes.Count > 0 Then
        Dim nodeItem As HtmlNode
        For Each nodeItem In inner_nodes
            
            special_get_string_from_collection = special_get_string_from_collection & nodeItem.to_string & Chr(10)
        Next nodeItem
    End If
End Function
