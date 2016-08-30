Attribute VB_Name = "Convert2HTMLModule"
Public Function fnConvert2HTML(r As Word.Range) As String

    Dim bldTagOn, itlTagOn, ulnTagOn, colTagOn As Boolean
    Dim i, chrCount As Integer
    Dim chrCol, chrLastCol, htmlTxt, htmlEnd As String

    bldTagOn = False
    itlTagOn = False
    ulnTagOn = False
    colTagOn = False
    chrCol = "NONE"
    'htmlTxt = "<html>"
    htmlTxt = ""
    chrCount = r.Characters.Count

    For i = 1 To chrCount
    htmlEnd = ""
        With r.Characters(i)
            If (.Font.Color) Then
                chrCol = fnGetCol(.Font.Color)
                If Not colTagOn Then
                    htmlTxt = htmlTxt & "<font color=#" & chrCol & ">"
                    colTagOn = True
                Else
                    If chrCol <> chrLastCol Then htmlTxt = htmlTxt & "</font><font color=#" & chrCol & ">"
                End If
            Else
                chrCol = "NONE"
                If colTagOn Then
                    htmlEnd = "</font>" & htmlEnd
                    'htmlTxt = htmlTxt & "</font>"
                    colTagOn = False
                End If
            End If
            chrLastCol = chrCol

            If .Font.Bold = True Then
                If Not bldTagOn Then
                    htmlTxt = htmlTxt & "<b>"
                    bldTagOn = True
                End If
            Else
                If bldTagOn Then
                    'htmlTxt = htmlTxt & "</b>"
                    htmlEnd = "</b>" & htmlEnd
                    bldTagOn = False
                End If
            End If

            If .Font.Italic = True Then
                If Not itlTagOn Then
                    htmlTxt = htmlTxt & "<i>"
                    itlTagOn = True
                End If
            Else
                If itlTagOn Then
                    'htmlTxt = htmlTxt & "</i>"
                    htmlEnd = "</i>" & htmlEnd
                    itlTagOn = False
                End If
            End If

            If .Font.Underline > 0 Then
                If Not ulnTagOn Then
                    htmlTxt = htmlTxt & "<u>"
                    ulnTagOn = True
                End If
            Else
                If ulnTagOn Then
                    'htmlTxt = htmlTxt & "</u>"
                    htmlEnd = "</u>" & htmlEnd
                    ulnTagOn = False
                End If
            End If

            If (Asc(.Text) = 10) Then
                htmlTxt = htmlTxt & htmlEnd & "<br>"
            Else
                htmlTxt = htmlTxt & htmlEnd & .Text
            End If

        End With
    Next

    If colTagOn Then
        htmlTxt = htmlTxt & "</font>"
        colTagOn = False
    End If
    If bldTagOn Then
        htmlTxt = htmlTxt & "</b>"
        bldTagOn = False
    End If
    If itlTagOn Then
        htmlTxt = htmlTxt & "</i>"
        itlTagOn = False
    End If
    If ulnTagOn Then
        htmlTxt = htmlTxt & "</u>"
        ulnTagOn = False
    End If
    'htmlTxt = htmlTxt & "</html>"
    fnConvert2HTML = htmlTxt
End Function

Public Function fnGetCol(strCol As String) As String
    Dim rVal, gVal, bVal As String
    strCol = Right("000000" & Hex(strCol), 6)
    bVal = Left(strCol, 2)
    gVal = Mid(strCol, 3, 2)
    rVal = Right(strCol, 2)
    fnGetCol = rVal & gVal & bVal
End Function
