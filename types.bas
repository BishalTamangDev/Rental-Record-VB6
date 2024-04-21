Attribute VB_Name = "Module1"
Type MainUser
    dob_year As String * 5
    dob_date As String * 2
    dob_month As String * 9
    username As String * 25
    password As String * 35
    pin_number As String * 5
    contact_number As String * 15
    security_question As String * 25
End Type

Type room_class
    bhk As Integer
    rent_amount As Single
    room_number As Integer
    room_occupied As Boolean
    tenant_fname As String * 10
    tenant_mname As String * 10
    tenant_lname As String * 10
    tenant_contact As String * 15
    service_provided As String * 57
    'rent_status As String * 7
End Type

Type add_room_display
    sn As String * 7
    bhk As String * 6
    rent As String * 14
    occupied As String * 11
    room_num As String * 10
    tenant_fname As String * 10
    tenant_mname As String * 10
    tenant_lname As String * 15
    tenant_contact As String * 18
    'rent_status As String * 13
End Type

Type tenant_class
    'rent_end As Date
    'rent_start As Date
    room_num As Integer
    last_name As String * 10
    first_name As String * 10
    middle_name As String * 10
    contact_num As String * 15
    citizenship As String * 15
    address_ward As String * 3
    
    rent_year(1) As Integer
    rent_month(1) As Integer
    
    address_district As String * 15
    address_municipality As String * 15
End Type

Type tenant_class_display
    serial As String * 6
    
    'rent_end As String * 22
    'rent_start As String * 22
    rent_year(1) As String * 5
    rent_month(1) As String * 17
    
    room_num As String * 10
    last_name As String * 11
    first_name As String * 10
    middle_name As String * 15
    contact_num As String * 15
    citizenship As String * 23
    address_ward As String * 6
    address_district As String * 15
    address_municipality As String * 15
End Type

Type issue_class
    room_num As Integer
    contact_num As String * 15
    reporter As String * 25
    issue_detail As String * 100
    issue_status As String * 9
    issue_reported_date As Date
    issue_solved_date As Date
End Type

Type display_issue_class
    serial As String * 6
    room_num As String * 10
    contact_num As String * 13
    reporter As String * 24
    issue_detail As String
    issue_status As String * 14
    issue_reported_date As String * 19
    issue_solved_date As String * 19
End Type

Type service
    service_year(1) As Integer
    service_month(1) As Integer
    fee(3) As Single
End Type

Type display_service
    serial As String * 6
    fee(3) As String * 12
    service_year(1) As String * 9
    service_month(1) As String * 10
End Type

Type electricity
    per_unit(6) As Single
    range_min(6) As Integer
    range_max(6) As Integer
    monthly_min(6) As Integer
    electricity_year(1) As Integer
    electricity_month(1) As Integer
End Type

Type display_electricity
    serial As String * 6
    range_min(6) As String * 8
    range_max(6) As String * 8
    monthly_min(6) As String * 11
    per_unit(6) As String * 6
    electricity_year(1) As String * 9
    electricity_month(1) As String * 10
End Type

Type payment_class
    room_num As Integer
    
    fname As String * 10
    mname As String * 10
    lname As String * 10
    
    year As Integer
    month As Integer
    
    payment_year As Integer
    payment_month As Integer
    payment_date As Integer
    
    contact As String * 15
    
    rent_amount As Integer
    elec_unit As Integer
    
    water_fee As Integer
    waste_fee As Integer
    security_fee As Integer
    internet_fee As Integer
    
    electricity_fee As Single
    
    total As Single
    
    is_paid As Boolean
End Type

Type display_payment_class
    serial As String * 6
    room_num As String * 6
    
    fname As String * 10
    mname As String * 10
    lname As String * 10
    
    year As String * 5
    month As String * 10
    
    payment_year As String * 5
    payment_month As String * 10
    payment_date As String * 4
    
    contact As String * 16
    rent_amount As String * 9
    elec_unit As String * 13
    water_fee As String * 10
    waste_fee As String * 12
    security_fee As String * 10
    internet_fee As String * 10
    electricity_fee As String * 13
    total As String * 8
    is_paid As String * 7
End Type
