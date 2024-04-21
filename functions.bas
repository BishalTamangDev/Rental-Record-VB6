Attribute VB_Name = "Module2"
Dim has_paid As Boolean

Dim pending_present As Boolean

Dim today_year As Integer
Dim today_month As Integer
Dim today_day As Integer

Dim service_2 As String * 57
Dim service_1 As String * 57

Dim counter_1 As Integer
Dim counter_2 As Integer
Dim counter_3 As Integer
Dim counter_4 As Integer
Dim counter_5 As Integer

Dim counter_1_1 As Integer
Dim counter_1_2 As Integer

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim n As Integer

Dim sign_up_data As MainUser

Dim temp_room As room_class
Dim existing_room As room_class

Dim temp_issue As issue_class

Dim temp_service As service

Dim temp_tenant As tenant_class
Dim temp_tenant_1 As tenant_class

Dim new_payment As payment_class
Dim temp_payment As payment_class
Dim existing_payment As payment_class

Dim display_payment As display_payment_class

Dim temp_elect As electricity

Public Function show_pending_payment(num As Integer, string_1 As String, string_2 As String, string_3 As String)
    'frame visibility
    Payment_Form.WindowState = Tenant_Form.WindowState
    Payment_Form.Visible = True
    Payment_Form.payment_detail_frame.Visible = True
    Tenant_Form.Visible = False
    
    'reset values of tenant form >> remove tenant
    Tenant_Form.remove_tenant_fr_combo_1.Text = "Room number"
    Tenant_Form.remove_tenant_fr_TextBox(0).Text = "First Name"
    Tenant_Form.remove_tenant_fr_TextBox(1).Text = "Middle Name"
    Tenant_Form.remove_tenant_fr_TextBox(2).Text = "Last Name"
    Tenant_Form.remove_tenant_fr_remove_command.Caption = "Remove"
    
    Payment_Form.payment_detail_TextBox.Text = ""
    Payment_Form.payment_filter_command.Caption = "Filter : On"
    Payment_Form.payment_filter_option(1).Value = True
    Payment_Form.payment_filter_check(3).Value = Checked
    Payment_Form.payment_filter_check(0).Value = Checked
    
    counter_2 = room_detail_count_function
    Open "RoomDetail.txt" For Random As #2 Len = 112
        For counter_1 = 1 To counter_2
            Get #2, counter_1, temp_room
            Payment_Form.payment_filter_combo(0).AddItem temp_room.room_number
        Next counter_1
    Close #2
    Payment_Form.payment_filter_combo(0).Text = "Room No"
    
    Payment_Form.payment_filter_combo(1).Clear
    For i = 2022 To (return_year + 1)
        Payment_Form.payment_filter_combo(1).AddItem i
    Next i
    Payment_Form.payment_filter_combo(1).Text = "Year"
    
    
    Payment_Form.payment_filter_combo(0).Text = num
    
    'filtering payment details
    counter_2 = payment_detail_count_function
    Open "PaymentDetail.txt" For Random As #1 Len = 79
        counter_3 = 1
        For counter_1 = 1 To counter_2 Step 1
            Get #1, counter_1, existing_payment
            display_payment.serial = counter_3
            display_payment.room_num = existing_payment.room_num
            display_payment.fname = existing_payment.fname
                
            If existing_payment.mname = "Null      " Then
                display_payment.mname = existing_payment.lname
                display_payment.lname = ""
            Else
                display_payment.mname = existing_payment.mname
                display_payment.lname = existing_payment.lname
            End If
                
            display_payment.contact = existing_payment.contact
            display_payment.year = existing_payment.year
            display_payment.rent_amount = existing_payment.rent_amount
            display_payment.water_fee = existing_payment.water_fee
            display_payment.waste_fee = existing_payment.waste_fee
            display_payment.security_fee = existing_payment.security_fee
            display_payment.internet_fee = existing_payment.internet_fee
                
            If existing_payment.total = 0# Then
                display_payment.payment_year = "-"
                display_payment.payment_month = "-"
                display_payment.payment_date = "-"
                display_payment.elec_unit = "-"
                display_payment.electricity_fee = "-"
                display_payment.total = "-"
            Else
                display_payment.payment_year = existing_payment.payment_year
                display_payment.payment_month = existing_payment.payment_month
                display_payment.payment_date = existing_payment.payment_date
                display_payment.elec_unit = existing_payment.elec_unit
                display_payment.electricity_fee = existing_payment.electricity_fee
                display_payment.total = existing_payment.total
            End If
                
            Call return_month_in_string(display_payment.month, existing_payment.month)
            Call return_month_in_string(display_payment.payment_month, existing_payment.payment_month)
            
            'filter unpaid rent details
            
            If existing_payment.room_num = num Then
                If existing_payment.fname = string_1 Then
                    If existing_payment.mname = string_2 Then
                        If existing_payment.lname = string_3 Then
                            If existing_payment.total = 0# Then
                                Call display_payment_detail_function(display_payment)
                                counter_3 = counter_3 + 1
                            End If
                        End If
                    End If
                End If
            End If
            Next counter_1
        Close #1
End Function


Public Function fill_payment_details()
    'displaying payment details
    Dim display_payment As display_payment_class
    Dim exist_pay As payment_class
    
    Payment_Form.payment_detail_TextBox.Text = ""
    
    counter_2 = payment_detail_count_function
    counter_1 = 1
    Open "PaymentDetail.txt" For Random As #1 Len = 79
        For counter_1 = 1 To counter_2 Step 1
            Get #1, counter_1, exist_pay
            
            display_payment.serial = counter_1
            display_payment.room_num = exist_pay.room_num
            display_payment.fname = exist_pay.fname
            
            If exist_pay.mname = "Null      " Then
                display_payment.mname = exist_pay.lname
                display_payment.lname = ""
            Else
                display_payment.mname = exist_pay.mname
                display_payment.lname = exist_pay.lname
            End If
            
            display_payment.contact = exist_pay.contact
            display_payment.year = exist_pay.year
            display_payment.rent_amount = exist_pay.rent_amount
            display_payment.water_fee = exist_pay.water_fee
            display_payment.waste_fee = exist_pay.waste_fee
            display_payment.security_fee = exist_pay.security_fee
            display_payment.internet_fee = exist_pay.internet_fee
            
            If exist_pay.elec_unit = 0 Then
                display_payment.payment_year = "-"
                display_payment.payment_date = "-"
                display_payment.elec_unit = "-"
                display_payment.electricity_fee = "-"
                display_payment.total = "-"
            Else
                display_payment.payment_year = exist_pay.payment_year
                display_payment.payment_date = exist_pay.payment_date
                display_payment.elec_unit = exist_pay.elec_unit
                display_payment.electricity_fee = exist_pay.electricity_fee
                display_payment.total = exist_pay.total
            End If
            
            Call return_month_in_string(display_payment.month, exist_pay.month)
            Call return_month_in_string(display_payment.payment_month, exist_pay.payment_month)
            
            Payment_Form.payment_detail_TextBox.Text = Payment_Form.payment_detail_TextBox.Text + display_payment.serial + display_payment.room_num + display_payment.fname + display_payment.mname + display_payment.lname + display_payment.contact
            Payment_Form.payment_detail_TextBox.Text = Payment_Form.payment_detail_TextBox.Text + display_payment.year + display_payment.month + display_payment.payment_year + display_payment.payment_month + display_payment.payment_date
            Payment_Form.payment_detail_TextBox.Text = Payment_Form.payment_detail_TextBox.Text + display_payment.rent_amount + display_payment.water_fee + display_payment.waste_fee + display_payment.security_fee + display_payment.elec_unit + display_payment.electricity_fee
            Payment_Form.payment_detail_TextBox.Text = Payment_Form.payment_detail_TextBox.Text + display_payment.internet_fee + display_payment.total + vbNewLine
        Next counter_1
    Close #1
End Function

'calculate electricity fee and return >> single value
Public Function calculate_electricity_fee(unit As Integer, year As Integer, month As Integer) As Single
    Dim fee As Single
    fee = 0#
    
    counter_1_2 = electricity_fee_count_function
    If counter_1_2 > 0 Then
        Open "ElectricityFee.txt" For Random As #6 Len = 78
            For counter_1_1 = 1 To counter_1_2 Step 1
                Get #6, counter_1_1, temp_elect
                
                If temp_elect.electricity_year(1) = 0 And temp_elect.electricity_month(1) = 0 Then
                    If year >= temp_elect.electricity_year(0) And month >= temp_elect.electricity_month(0) Then
                        fee = electricity_fee_function(unit, temp_elect)
                        counter_1_1 = counter_1_2 + 1
                    End If
                Else
                    If temp_elect.electricity_year(0) = temp_elect.electricity_year(1) And temp_elect.electricity_year(0) = year Then
                        If month >= temp_elect.electricity_month(0) And month < temp_elect.electricity_month(1) Then
                            fee = electricity_fee_function(unit, temp_elect)
                            counter_1_1 = counter_1_2 + 1
                        End If
                    
                    ElseIf temp_elect.electricity_year(0) = temp_elect.electricity_year(1) And temp_elect.electricity_year(0) <> year Then
                    'discard this case
                    
                    ElseIf temp_elect.electricity_year(0) <> temp_elect.electricity_year(1) Then
                        If year >= temp_elect.electricity_year(0) And year <= temp_elect.electricity_year(1) Then
                            fee = electricity_fee_function(unit, temp_elect)
                            counter_1_1 = counter_1_2 + 1
                        End If
                    End If
                End If
            Next counter_1_1
        Close #6
    Else
        fee = 0#
    End If
    calculate_electricity_fee = fee
End Function


Public Function electricity_fee_function(unit As Integer, temp_elect As electricity) As Single
    Dim total As Single
   
    total = 0#
    temp = unit
    
    For i = 0 To 5 Step 1
        If (unit) > temp_elect.range_max(i) Then
            If i = 0 Then
                total = total + (temp_elect.range_max(i) - temp_elect.range_min(i)) * temp_elect.per_unit(i)
                temp = temp - temp_elect.range_max(i)
            Else
                total = total + (temp_elect.range_max(i) - temp_elect.range_min(i) + 1) * temp_elect.per_unit(i)
                temp = temp - temp_elect.range_max(i) + temp_elect.range_min(i) - 1
            End If
        Else
            If i = 0 Then
                total = total + temp_elect.monthly_min(i)
            Else
                total = total + temp * temp_elect.per_unit(i) + temp_elect.monthly_min(i)
            End If
            i = 6
        End If
    Next i
    electricity_fee_function = total
End Function


Public Function get_pending_detail(new_pay As payment_class, num As Integer, string_1 As String, string_2 As String, string_3 As String)
    count_2 = payment_detail_count_function
    Open "PaymentDetail.txt" For Random As #5 Len = 79
        For count_1 = 1 To count_2 Step 1
            Get #5, count_1, new_pay
            If new_pay.room_num = num Then 'room number
                If new_pay.fname = string_1 And new_pay.mname = string_2 And new_pay.lname = string_3 Then
                    If new_pay.total = 0# Then
                        count_1 = count_2 + 1
                    End If
                End If
            End If
        Next count_1
    Close #5
End Function


Public Function update_payment_detail()
    counter_2 = room_detail_count_function
    If counter_2 > 0 Then
        Open "RoomDetail.txt" For Random As #1 Len = 112
            For counter_1 = 1 To counter_2
                Get #1, counter_1, temp_room
                If temp_room.room_occupied = True Then
                    m = return_rent_year(temp_room.room_number, temp_room.tenant_fname, temp_room.tenant_mname, temp_room.tenant_lname)
                    n = return_rent_month(temp_room.room_number, temp_room.tenant_fname, temp_room.tenant_mname, temp_room.tenant_lname)
                    
                    'check if rent is to be collected or not >> based on the date
                    has_paid = False
                    If return_year = m Then
                        If return_month > n Then
                            If return_month <> 1 Then
                                has_paid = payment_available_check(temp_room.room_number, return_year, return_month - 1, temp_room.tenant_fname, temp_room.tenant_mname, temp_room.tenant_lname)
                            Else
                                has_paid = payment_available_check(temp_room.room_number, return_year - 1, return_month, temp_room.tenant_fname, temp_room.tenant_mname, temp_room.tenant_lname)
                            End If
                        ElseIf return_month <= n Then
                            has_paid = True
                        End If
                    ElseIf return_year > m Then
                        has_paid = payment_available_check(temp_room.room_number, return_year, return_month, temp_room.tenant_fname, temp_room.tenant_mname, temp_room.tenant_lname)
                    End If
                                       
                    new_payment.room_num = temp_room.room_number
                    new_payment.fname = temp_room.tenant_fname
                    new_payment.mname = temp_room.tenant_mname
                    new_payment.lname = temp_room.tenant_lname
                    new_payment.contact = temp_room.tenant_contact
                    new_payment.rent_amount = temp_room.rent_amount
                                                
                    If return_month > 1 Then
                        new_payment.year = return_year
                        new_payment.month = return_month - 1
                    Else
                        new_payment.year = return_year - 1
                        new_payment.month = return_month
                    End If
                    
                    new_payment.payment_year = 0
                    new_payment.payment_month = 0
                    new_payment.payment_date = 0
                    new_payment.electricity_fee = 0
                    new_payment.total = 0
                    new_payment.is_paid = False
                    new_payment.elec_unit = 0
                
                    Call get_service_fee(return_year, return_month, new_payment.water_fee, new_payment.waste_fee, new_payment.security_fee, new_payment.internet_fee, temp_room.service_provided)
                    
                    service_2 = "Water, Waste Management, Electricity, Security, Internet"
                    
                    If temp_room.service_provided = service_2 Then
                        new_payment.internet_fee = new_payment.internet_fee
                    Else
                        new_payment.internet_fee = 0#
                    End If
                       
                    If has_paid = False Then 'write detail in a file
                        counter_3 = payment_detail_count_function
                        Open "PaymentDetail.txt" For Random As #2 Len = 79
                            If counter_3 = 0 Then
                                Put #2, 1, new_payment
                            Else
                                Put #2, counter_3 + 1, new_payment
                            End If
                        Close #2
                    End If
                End If
            Next counter_1
        Close #1
    End If
End Function


Public Function get_service_fee(year As Integer, month As Integer, rate_1 As Integer, rate_2 As Integer, rate_3 As Integer, rate_4 As Integer, service_3 As String)
    rate_1 = 0
    rate_2 = 0
    rate_3 = 0
    rate_4 = 0

    service_1 = "Water, Waste Management, Electricity, Security, Internet"
        
    counter_1_2 = service_count_function
    If service_count_function > 0 Then
        Open "ServiceFee.txt" For Random As #6 Len = 24
            For counter_1_1 = 1 To counter_1_2 Step 1
                Get #6, counter_1_1, temp_service
                
                'recently added service fee >> service file
                If temp_service.service_year(1) = 0 And temp_service.service_month(1) = 0 Then
                    If year >= temp_service.service_year(0) And return_month >= temp_service.service_month(0) Then
                        rate_1 = temp_service.fee(0)
                        rate_2 = temp_service.fee(1)
                        rate_3 = temp_service.fee(2)
                        rate_4 = temp_service.fee(3)
                        counter_1_1 = counter_1_2 + 1
                    End If
                Else
                    'same year (Current year and file date)
                    If temp_service.service_year(0) = temp_service.service_year(1) And temp_service.service_year(0) = year Then
                        If month >= temp_service.service_month(0) And (month - 1) < temp_service.service_month(1) Then
                            rate_1 = temp_service.fee(0)
                            rate_2 = temp_service.fee(1)
                            rate_3 = temp_service.fee(2)
                            rate_4 = temp_service.fee(3)
                            counter_1_1 = counter_1_2 + 1
                        End If
                        
                    'same year but less than current year >> discard this case
                    ElseIf temp_service.service_year(0) = temp_service.service_year(1) And temp_service.service_year(0) <> year Then
                        rate_1 = rate_1
                    
                    'different year
                    ElseIf temp_service.service_year(0) <> temp_service.service_year(1) Then
                        If year >= temp_service.service_year(0) And year <= temp_service.service_year(1) Then
                            rate_1 = temp_service.fee(0)
                            rate_2 = temp_service.fee(1)
                            rate_3 = temp_service.fee(2)
                            rate_4 = temp_service.fee(3)
                            counter_1_1 = counter_1_2 + 1
                        End If
                    End If
                End If
            Next counter_1_1
        Close #6
    Else
        rate_1 = 0
        rate_2 = 0
        rate_3 = 0
        rate_4 = 0
    End If
End Function


Public Function payment_available_check(num As Integer, year, month, string_1 As String, string_2 As String, string_3 As String) As Boolean
    Dim status As Boolean
    status = False
    
    j = payment_detail_count_function
    If counter_1_2 > 0 Then
        Open "PaymentDetail.txt" For Random As #5 Len = 79
            For i = 1 To j
                Get #5, i, temp_payment
                If temp_payment.room_num = num Then 'same room number
                    If temp_payment.year = year And temp_payment.month = month Then    'same date
                        If temp_payment.fname = string_1 And temp_payment.mname = string_2 And temp_payment.lname = string_3 Then
                            status = True
                        End If
                    End If
                End If
            Next i
        Close #5
    Else
        status = True
    End If
    payment_available_check = status
End Function


Public Function pending_payment_check(num As Integer, string_1 As String, string_2 As String, string_3 As String) As Boolean
    Dim status As Boolean
    status = False
    j = payment_detail_count_function
    
    If j > 0 Then
        Open "PaymentDetail.txt" For Random As #5 Len = 79
            For i = 1 To j Step 1
                Get #5, i, temp_payment
                If temp_payment.room_num = num Then 'same room number
                    If temp_payment.fname = string_1 Then
                        If temp_payment.mname = string_2 Then
                            If temp_payment.lname = string_3 Then
                                If temp_payment.is_paid = False Then
                                    status = True
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        Close #5
    Else
        status = True
    End If
    pending_payment_check = status
End Function


Public Function return_rent_year(num_1 As Integer, string_1 As String, string_2 As String, string_3 As String) As Integer
    Dim num_2 As Integer
    counter_1_2 = tenant_detail_count_function
    
    Open "TenantDetail.txt" For Random As #5 Len = 103
        For counter_1_1 = 1 To counter_1_2
            Get #5, counter_1_1, temp_tenant
            If temp_tenant.room_num = num_1 Then
                If temp_tenant.first_name = string_1 Then
                    If temp_tenant.middle_name = string_2 Then
                        If temp_tenant.last_name = string_3 Then
                            num_2 = temp_tenant.rent_year(0)
                        End If
                    End If
                End If
            End If
        Next counter_1_1
    Close #5
    return_rent_year = num_2
End Function


Public Function return_rent_month(num_1 As Integer, string_1 As String, string_2 As String, string_3 As String) As Integer
    Dim num_2 As Integer
    counter_1_2 = tenant_detail_count_function
    
    Open "TenantDetail.txt" For Random As #5 Len = 103
        For counter_1_1 = 1 To counter_1_2
            Get #5, counter_1_1, temp_tenant
            If temp_tenant.room_num = num_1 Then
                If temp_tenant.first_name = string_1 Then
                    If temp_tenant.middle_name = string_2 Then
                        If temp_tenant.last_name = string_3 Then
                            num_2 = temp_tenant.rent_month(0)
                        End If
                    End If
                End If
            End If
        Next counter_1_1
    Close #5
    return_rent_month = num_2
End Function

Sub Main()
    Call get_year
    Call get_month
    Call get_day
    
    Call update_payment_detail
    
    'form size
    Main_Form.WindowState = 2
    
    'main form >> frame visibility
    Main_Form.LogInFrame.Visible = False
    Main_Form.SignUpFrame.Visible = False
    Main_Form.ForgotPasswordFrame.Visible = False
    Main_Form.main_menu_frame.Visible = False
    Main_Form.room_frame.Visible = False
    Main_Form.add_room_frame.Visible = False
    Main_Form.remove_room_frame.Visible = False
    Main_Form.edit_room_frame.Visible = False
    Main_Form.edit_room_frame_2.Visible = False
    Main_Form.Reset_Frame.Visible = False
    
    'tenant form >> frame visibility
    Tenant_Form.tenant_frame.Visible = False
    Tenant_Form.add_tenant_frame.Visible = False
    Tenant_Form.edit_tenant_frame.Visible = False
    Tenant_Form.remove_tenant_frame.Visible = False
    
    'payment form >> frame visibility
    Payment_Form.service_det_frame.Visible = False
    Payment_Form.electricity_fee_frame.Visible = False
    Payment_Form.service_add_fr.Visible = False
    Payment_Form.payment_frame.Visible = False
    Payment_Form.payment_detail_frame.Visible = False
    Payment_Form.payment_confirm_fr.Visible = False
    
    'issue form >> frame visibility
    Issue_Form.issue_report_frame.Visible = False
    Issue_Form.issue_detail_frame.Visible = False
    
    'payment form >> frame visibility
    Payment_Form.payment_frame.Visible = False
    Payment_Form.service_add_fr.Visible = False
    Payment_Form.service_add_fr.Visible = False
    Payment_Form.service_det_frame.Visible = False
    Payment_Form.payment_detail_frame.Visible = False
    
    'main form >> option frame visibility
    Main_Form.room_option_frame.Visible = False
    Main_Form.tenant_option_frame.Visible = False
    Main_Form.payment_option_frame.Visible = False
    Main_Form.issue_option_frame.Visible = False
        
    'default value assignment
    option_count = 1
    initial_date = 0
    Main_Form.login_pw_hide.Value = Checked
    Main_Form.signup_pw_hide.Value = Checked
    Main_Form.LogInFrame_Password_TextBox.PasswordChar = ""
    Main_Form.SignUpFrame_Password_TextBox.PasswordChar = ""
    Main_Form.SignUpFrame_PasswordConfirmation_TextBox.PasswordChar = ""
    
    'signup >> add year
    Main_Form.SignUpFrame_dob_Combo(0).Clear
    Main_Form.ForgotPasswordFrame_dob_Combo(0).Clear
    
    For i = Format(Now, "yyyy") To Format(Now, "yyyy") - 80 Step -1
        Main_Form.SignUpFrame_dob_Combo(0).AddItem i
        Main_Form.ForgotPasswordFrame_dob_Combo(0).AddItem i
    Next i
    
    Main_Form.SignUpFrame_dob_Combo(0).Text = "Year"
    Main_Form.ForgotPasswordFrame_dob_Combo(0).Text = "Year"
    
    'signup_pw_hide.Value = Unchecked
    LogInFrame_PictureTimer = 1
    SignUpFrame_PictureTimer = 1
    ForgotPasswordFrame_PictureTimer = 1
    Reset_Frame_PictureTimer = 1
    
    Main_Form.LogInFrame_Timer.Enabled = True
    Main_Form.SignUpFrame_Timer.Enabled = True
    Main_Form.ForgotPasswordFrame_Timer.Enabled = False
    Main_Form.reset_frame_timer.Enabled = False
    
    'service_1 = "Water, Waste Management, Electricity, Security"
    'service_2 = "Water, Waste Management, Electricity, Security, Internet"

    'menu image visibility
    Main_Form.option_image(0).Visible = True
    Main_Form.option_image(1).Visible = False
    Main_Form.option_image(2).Visible = False
    Main_Form.option_image(3).Visible = False
    Main_Form.option_image(4).Visible = False

    'create mandatory files
    Main_Form.Visible = True
    Call create_file_function
    
    'sign up detail presence
    On Error GoTo ErrorHandler1
    Open "AdminData.txt" For Random As #1 Len = 121
        Get #1, 1, sign_up_data
        If sign_up_data.dob_date = 0 Then GoTo ErrorHandler1
    Close #1
    
    'change
    Main_Form.LogInFrame.Visible = False
    Main_Form.main_menu_frame.Visible = True
    Main_Form.SignUpFrame.Visible = False
    Exit Sub
    
ErrorHandler1:
    Close #1
    
    Main_Form.SignUpFrame.Visible = True
    Main_Form.LogInFrame.Visible = False
End Sub


'create a mandatory files
Public Function create_file_function()
    Open "RoomDetail.txt" For Random As #1
    Open "TenantDetail.txt" For Random As #2
    Open "PaymentDetail.txt" For Random As #3
    Open "IssueDetail.txt" For Random As #4
    Open "ServiceFee.txt" For Random As #5
    Open "ElectricityFee.txt" For Random As #6
    Close #1
    Close #2
    Close #3
    Close #4
    Close #5
    Close #6
End Function


'count number of existing room details
Public Function room_detail_count_function() As Integer
    counter_1_1 = 1
    Open "RoomDetail.txt" For Random As #5 Len = 112
        While EOF(5) = False
            Get #5, counter_1_1, temp_room
            counter_1_1 = counter_1_1 + 1
        Wend
    Close #5
    counter_1_1 = counter_1_1 - 2
    room_detail_count_function = counter_1_1
End Function

'count number of existing room details
Public Function tenant_detail_count_function() As Integer
    counter_1_1 = 1
    Open "TenantDetail.txt" For Random As #5 Len = 103
        While EOF(5) = False
            Get #5, counter_1_1, temp_tenant
            counter_1_1 = counter_1_1 + 1
        Wend
    Close #5
    counter_1_1 = counter_1_1 - 2
    tenant_detail_count_function = counter_1_1
End Function

'count number of existing issues details
Public Function issue_count_function() As Integer
    counter_1_1 = 1
    Open "IssueDetail.txt" For Random As #5 Len = 167
        While EOF(5) = False
            Get #5, counter_1_1, temp_issue
            counter_1_1 = counter_1_1 + 1
        Wend
    Close #5
    counter_1_1 = counter_1_1 - 2
    issue_count_function = counter_1_1
End Function

'count number of existing fee details
Public Function service_count_function() As Integer
    counter_1_1 = 1
    Open "ServiceFee.txt" For Random As #5 Len = 24
        While EOF(5) = False
            Get #5, counter_1_1, temp_service
            counter_1_1 = counter_1_1 + 1
        Wend
    Close #5
    counter_1_1 = counter_1_1 - 2
    service_count_function = counter_1_1
End Function

'count number of existing electricity details
Public Function electricity_fee_count_function() As Integer
    counter_1_1 = 1
    Open "ElectricityFee.txt" For Random As #5 Len = 78
        While EOF(5) = False
            Get #5, counter_1_1, temp_electricity
            counter_1_1 = counter_1_1 + 1
        Wend
    Close #5
    counter_1_1 = counter_1_1 - 2
    electricity_fee_count_function = counter_1_1
End Function

Public Function payment_detail_count_function()
    counter_1_1 = 1
    Open "PaymentDetail.txt" For Random As #5 Len = 79
        While EOF(5) = False
            Get #5, counter_1_1, temp_payment
            counter_1_1 = counter_1_1 + 1
        Wend
    Close #5
    counter_1_1 = counter_1_1 - 2
    payment_detail_count_function = counter_1_1
End Function

'returns the month in string for a month provided in number
Public Function return_month_in_string(temp_string As String, temp_num As Integer)
    If temp_num = 1 Then
        temp_string = "January"
    ElseIf temp_num = 2 Then
        temp_string = "February"
    ElseIf temp_num = 3 Then
        temp_string = "March"
    ElseIf temp_num = 4 Then
        temp_string = "April"
    ElseIf temp_num = 5 Then
        temp_string = "May"
    ElseIf temp_num = 6 Then
        temp_string = "June"
    ElseIf temp_num = 7 Then
        temp_string = "July"
    ElseIf temp_num = 8 Then
        temp_string = "August"
    ElseIf temp_num = 9 Then
        temp_string = "September"
    ElseIf temp_num = 10 Then
        temp_string = "October"
    ElseIf temp_num = 11 Then
        temp_string = "November"
    ElseIf temp_num = 12 Then
        temp_string = "December"
    Else
        temp_string = "-"
    End If
End Function

'returns the integer value for the provided month
Public Function set_month_in_integer(temp_string As String) As Integer
    Dim temp_integer As Integer
    If temp_string = "January" Then
        temp_integer = 1
    ElseIf temp_string = "February" Then
        temp_integer = 2
    ElseIf temp_string = "March" Then
        temp_integer = 3
    ElseIf temp_string = "April" Then
        temp_integer = 4
    ElseIf temp_string = "May" Then
        temp_integer = 5
    ElseIf temp_string = "June" Then
        temp_integer = 6
    ElseIf temp_string = "July" Then
        temp_integer = 7
    ElseIf temp_string = "August" Then
        temp_integer = 8
    ElseIf temp_string = "September" Then
        temp_integer = 9
    ElseIf temp_string = "October" Then
        temp_integer = 10
    ElseIf temp_string = "November" Then
        temp_integer = 11
    ElseIf temp_string = "December" Then
        temp_integer = 12
    End If
    set_month_in_integer = temp_integer
End Function

Public Function display_payment_detail_function(dis_pay As display_payment_class)
    Payment_Form.payment_detail_TextBox.Text = Payment_Form.payment_detail_TextBox.Text + dis_pay.serial + dis_pay.room_num + dis_pay.fname + dis_pay.mname + dis_pay.lname
    Payment_Form.payment_detail_TextBox.Text = Payment_Form.payment_detail_TextBox.Text + dis_pay.contact + dis_pay.year + dis_pay.month
    Payment_Form.payment_detail_TextBox.Text = Payment_Form.payment_detail_TextBox.Text + dis_pay.payment_year + dis_pay.payment_month + dis_pay.payment_date + dis_pay.rent_amount
    Payment_Form.payment_detail_TextBox.Text = Payment_Form.payment_detail_TextBox.Text + dis_pay.water_fee + dis_pay.waste_fee + dis_pay.security_fee + dis_pay.elec_unit + dis_pay.electricity_fee
    Payment_Form.payment_detail_TextBox.Text = Payment_Form.payment_detail_TextBox.Text + dis_pay.internet_fee + dis_pay.total + vbNewLine
End Function

Public Function get_year()
    today_year = Val(InputBox("Enter today's year : "))
    'today_year = Format(Now, "yyyy")
End Function

Public Function get_month()
    today_month = Val(InputBox("Enter today's month : "))
    'today_month = Format(Now, "mm")
End Function

Public Function get_day()
    today_day = Val(InputBox("Enter today's date : "))
    'today_date = Format(Now, "dd")
End Function

Public Function return_day() As Integer
    return_day = today_day
End Function

Public Function return_year() As Integer
    return_year = today_year
End Function

Public Function return_month() As Integer
    return_month = today_month
End Function
