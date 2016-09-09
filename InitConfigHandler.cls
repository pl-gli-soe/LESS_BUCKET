VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitConfigHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


' common
Public work_with_visible_data As Boolean
Public send_only_table As Boolean

' atts
Public att_labels As Boolean
Public att_blue_tag As Boolean
Public att_ppqp_label As Boolean
Public att_survey As Boolean
Public att_del_conf As Boolean
Public att_guide As Boolean
Public att_part_order_conf As Boolean
Public att_material_order As Boolean
Public att_ppqp_action_plan As Boolean


' mail content
Public mail_general_info As Boolean
Public mail_yellow_pic As Boolean
Public mail_order_table As Boolean
Public mail_rest_of_the_mail As Boolean

' radio
Public radio_material_order_as_excel As Boolean
Public radio_material_order_as_html As Boolean


Public Sub set_stanard()

    With Me
        .att_blue_tag = False
        .att_del_conf = True
        .att_guide = True
        .att_labels = True
        .att_material_order = True
        .att_part_order_conf = False
        .att_ppqp_action_plan = False
        .att_ppqp_label = False
        .att_survey = True
        .mail_general_info = True
        .mail_order_table = True
        .mail_rest_of_the_mail = True
        .mail_yellow_pic = True
        .work_with_visible_data = True
        .send_only_table = False
        .radio_material_order_as_excel = True
        .radio_material_order_as_html = False
    End With
End Sub

Public Sub set_biw()

    With Me
        .att_blue_tag = True
        .att_del_conf = True
        .att_guide = True
        .att_labels = True
        .att_material_order = True
        .att_part_order_conf = False
        .att_ppqp_action_plan = True
        .att_ppqp_label = True
        .att_survey = True
        .mail_general_info = True
        .mail_order_table = True
        .mail_rest_of_the_mail = True
        .mail_yellow_pic = True
        .work_with_visible_data = True
        .send_only_table = False
        .radio_material_order_as_excel = True
        .radio_material_order_as_html = False
    End With
End Sub

Public Sub set_psa()

    With Me
        .att_blue_tag = False
        .att_del_conf = True
        .att_guide = True
        .att_labels = True
        .att_material_order = True
        .att_part_order_conf = True
        .att_ppqp_action_plan = False
        .att_ppqp_label = False
        .att_survey = True
        .mail_general_info = True
        .mail_order_table = True
        .mail_rest_of_the_mail = True
        .mail_yellow_pic = True
        .work_with_visible_data = True
        .send_only_table = False
        .radio_material_order_as_excel = True
        .radio_material_order_as_html = False
    End With
End Sub

Public Sub set_table_only()

    With Me
        .att_blue_tag = False
        .att_del_conf = False
        .att_guide = False
        .att_labels = False
        .att_material_order = True
        .att_part_order_conf = False
        .att_ppqp_action_plan = False
        .att_ppqp_label = False
        .att_survey = False
        .mail_general_info = False
        .mail_order_table = True
        .mail_rest_of_the_mail = False
        .mail_yellow_pic = False
        .work_with_visible_data = False
        .send_only_table = True
    End With
End Sub

Public Sub adjust_checkboxes_and_radios(f As WybierzPlikForm)

    With Me
        f.CheckBoxBlueTag = .att_blue_tag
        f.CheckBoxDelConf = .att_del_conf
        f.CheckBoxGeneralInfo1 = .mail_general_info
        f.CheckBoxGuide = .att_guide
        f.CheckBoxLabels = .att_labels
        f.CheckBoxMaterialOrder = .att_material_order
        f.CheckBoxOrder = .mail_order_table
        f.CheckBoxPartOrderConf = .att_part_order_conf
        f.CheckBoxPPQP_BIW_label = .att_ppqp_label
        f.CheckBoxPPQP1Quality = .att_ppqp_action_plan
        f.CheckBoxRest = .mail_rest_of_the_mail
        f.CheckBoxSurvey = .att_survey
        f.CheckBoxTableOnly = .send_only_table
        f.CheckBoxVisible = .work_with_visible_data
        f.CheckBoxYellowPic = .mail_yellow_pic
        
        f.OptionButtonMaterialOrderTableAsExcel = .radio_material_order_as_excel
        f.OptionButtonMaterialOrderTableAsHTML = .radio_material_order_as_html
    End With

End Sub


Public Sub adjust_properties(f As WybierzPlikForm)

    With Me
        .att_blue_tag = f.CheckBoxBlueTag
        .att_del_conf = f.CheckBoxDelConf
        .mail_general_info = f.CheckBoxGeneralInfo1
        .att_guide = f.CheckBoxGuide
        .att_labels = f.CheckBoxLabels
        .att_material_order = f.CheckBoxMaterialOrder
        .mail_order_table = f.CheckBoxOrder
        .att_part_order_conf = f.CheckBoxPartOrderConf
        .att_ppqp_label = f.CheckBoxPPQP_BIW_label
        .att_ppqp_action_plan = f.CheckBoxPPQP1Quality
        .mail_rest_of_the_mail = f.CheckBoxRest
        .att_survey = f.CheckBoxSurvey
        .send_only_table = f.CheckBoxTableOnly
        .work_with_visible_data = f.CheckBoxVisible
        .mail_yellow_pic = f.CheckBoxYellowPic
        
        .radio_material_order_as_excel = f.OptionButtonMaterialOrderTableAsExcel
        .radio_material_order_as_html = f.OptionButtonMaterialOrderTableAsHTML
    End With
End Sub

