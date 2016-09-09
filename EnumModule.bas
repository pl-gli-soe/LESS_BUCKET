Attribute VB_Name = "EnumModule"
'source:
' http://www.rondebruin.nl/mac/mac020.htm
'
'These are the main file formats in Excel 2007-2013:
'
'51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
'52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)
'50 = xlExcel12 (Excel Binary Workbook in 2007-2013 with or without macro's, xlsb)
'56 = xlExcel8 (97-2003 format in Excel 2007-2013, xls)

Public Enum E_ACTION_TYPE
    E_ADD = 1
    E_EDIT
    E_REMOVE
    E_VIEW
End Enum


Public Enum E_FILE_FORMAT
    '51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
        xlOpenXMLWorkbook = 51
    '52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)
        xlOpenXMLWorkbookMacroEnabled = 52
    '50 = xlExcel12 (Excel Binary Workbook in 2007-2013 with or without macro's, xlsb)
        xlExcel12 = 50
    '56 = xlExcel8 (97-2003 format in Excel 2007-2013, xls)
        xlExcel8 = 56
End Enum


Public Enum E_MASTER_MANDATORY_COLUMNS
    pn = 1
    Alternative_PN
    PN_Name
    GPDS_PN_Name
    duns
    Supplier_Name
    country_code
    MGO_code
    Responsibility
    fup_code
    SQ
    ppap_status
    SQ_Comments
    MRD1_QTY
    MRD2_QTY
    Total_QTY
    ADD_to_T_slash_D
    MRD1_Ordered_date
    MRD1_Ordered_QTY
    MRD1_Ordered_STATUS
    MRD1_confirmed_qty
    MRD1_confirmed_qty_dot__Status
    MRD1_Total_PUS_STATUS
    MRD2_Ordered_date
    MRD2_Ordered_QTY
    MRD2_Ordered_STATUS
    MRD2_confirmed_qty
    MRD2_confirmed_qty_dot__Status
    MRD2_Total_PUS_STATUS
    Delivery_confirmation
    First_Confirmed_PUS_Date
    Delivery_reconfirmation
    Total_PUS_QTY
    Total_PUS_STATUS
    Comments
    Bottleneck
    Future_Osea
    DRE
    EDI_Received
    Capacity
    BLANK2
    BLANK3
    BLANK4
End Enum


Public Enum E_DOCS
    DOC_GUIDE = 1
    DOC_DEL_CONF
    DOC_PART_ORDER_CONF
    DOC_PPQP_1_Quality_Warrant_AND_PPQP_2_Corrective_Action_Plan
End Enum
