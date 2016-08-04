VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WybierzPlikForm 
   Caption         =   "Wybierz Plik typu Wizard"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3810
   OleObjectBlob   =   "WybierzPlikForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WybierzPlikForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private wh As WizardHandler

Private Sub BtnSubmit_Click()
    inner_run
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    inner_run

End Sub


Private Sub inner_run()


    If Len(CStr(Me.TextBox1)) = 2 Then

        hide
        
        
        Set wh = New WizardHandler
        
        If Me.ListBox1.ListCount > 0 Then
            If Me.ListBox1.Value <> "" Then
                wh.nazwa_aktywnego_pliku = Me.ListBox1.Value
            Else
                MsgBox "dla jakiego pliku chcesz zrobic ordery?"
                MsgBox "koncze z Toba wspolprace"
                Exit Sub
            End If
            
        Else
            MsgBox "nie ma czego wybrac!"
        End If
        
        
        
        If wh.nazwa_aktywnego_pliku <> "" Then
            ' lecimy dalej z logika
            ' textbox1 fu code
            ' combobox1.value - z ktorej kolumny robimy order
            '.ComboBox1.AddItem "MRD1 Qty"
            '.ComboBox1.AddItem "MRD2 Qty"
            '.ComboBox1.AddItem "Total Qty"
            '
            '.ComboBox1.AddItem "MRD1 Ordered Qty"
            '.ComboBox1.AddItem "MRD2 Ordered Qty"
            '
            '.ComboBox1.Value = "MRD1 Ordered Qty"
            
            Dim ee As E_MASTER_MANDATORY_COLUMNS
            
            
            If Me.ComboBox1.Value = "MRD1 Qty" Then
                ee = MRD1_QTY
            ElseIf Me.ComboBox1.Value = "MRD2 Qty" Then
                ee = MRD2_QTY
            ElseIf Me.ComboBox1.Value = "Total Qty" Then
                ee = Total_QTY
            ElseIf Me.ComboBox1.Value = "MRD1 Ordered Qty" Then
                ee = MRD1_Ordered_QTY
            ElseIf Me.ComboBox1.Value = "MRD2 Ordered Qty" Then
                ee = MRD2_Ordered_QTY
            End If
                
            wh.uruchom_dla_danych_funkcjonalnosc_w_wybranym_wizardzie ee, CStr(Me.TextBox1)
        End If
        
        
        Set wh = Nothing
    Else
        MsgBox "dla kogo chcesz zrobic ordery?"
    End If
End Sub

