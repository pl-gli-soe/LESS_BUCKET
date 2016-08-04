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


Public Sub init(nm As String, m As String, Optional ajdi As String, Optional stajl As String, Optional body As String)

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
    
    tmp = CStr(pre) & "<br/>" & _
        str_body & "<br/>" & _
        get_string_from_collection() & "<br/>" & CStr(post)
    
    to_string = CStr(tmp)
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
