Attribute VB_Name = "MainModule"
Global Const MASTER_SH_NAME = "MASTER"
Global Const DETAILS_SH_NAME = "DETAILS"

Public Sub main()

    
    With o.WybierzPlikForm
        .ListBox1.Clear
        .ComboBox1.Clear
        
        .ComboBox1.AddItem "MRD1 Qty"
        .ComboBox1.AddItem "MRD2 Qty"
        .ComboBox1.AddItem "Total Qty"
        
        .ComboBox1.AddItem "MRD1 Ordered Qty"
        .ComboBox1.AddItem "MRD2 Ordered Qty"
        
        .ComboBox1.Value = "MRD1 Ordered Qty"
        
        For Each s In Workbooks
            .ListBox1.AddItem CStr(s.name)
        Next s
        
        .Show
    End With
End Sub
