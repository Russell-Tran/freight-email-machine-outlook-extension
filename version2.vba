Sub newwindowocean()


    Dim OutMail As MailItem

    Set OutMail = Application.CreateItem(0)
    
    Dim company_name, po_number, personal_name, vessel_number, departing_city, departing_date, arriving_city, arriving_date, qty_name As String
    company_name = InputBox("Company Name?")
    po_number = InputBox("PO #?")
    personal_name = InputBox("Personal Name? (First name)")
    vessel_name = InputBox("Vessel Name?")
    departing_city = InputBox("Departing City?")
    departing_date = InputBox("Departing Date? (ETD)")
    arriving_city = InputBox("Arriving City?")
    arriving_date = InputBox("Arriving Date? (ETA)")
    qty_name = InputBox("Qty?")
    
    Dim subject As String
    subject = "PO to " & company_name & " - " & "PO# " & po_number
    
    
    Dim message As String
    message = "Good morning, " & personal_name
    message = message & vbNewLine & vbNewLine
    
    message = message & "We wanted to let you know that this PO# " & po_number & " has been booked as follows. "
    message = message & "It currently has an ETA into " & arriving_city & " of " & arriving_date & ". "
    message = message & "Once the documents are available, could you please send them to us? "
    message = message & vbNewLine & vbNewLine
    
    message = message
    message = message & "VESSEL : " & vessel_name & vbNewLine
    message = message & "ETD " & departing_city & " : " & departing_date & vbNewLine
    message = message & "ETA " & arriving_city & " : " & arriving_date & vbNewLine
    message = message & "Q'ty : " & qty_name & vbNewLine
    message = message

    On Error Resume Next
    With OutMail
        .To = ""
        .CC = "siobhansweeney@usffcl.com; kristiponce@usffcl.com"
        .BCC = ""
        .subject = .subject & subject
        .Body = .Body & message
        .Display
        If MsgBox("Send it?", vbYesNo) = vbYes Then .Send
    End With

    Set OutMail = Nothing
    
End Sub


Sub newwindowair()

    
    Dim OutMail As MailItem
    Set OutMail = Application.CreateItem(0)
    
    Dim company_name, po_number, personal_name, vessel_number, departing_city, departing_date, arriving_city, arriving_date, qty_name As String
    company_name = InputBox("Company Name?")
    po_number = InputBox("HWAB #?")
    personal_name = InputBox("Personal Name? (First name)")
    departing_city = InputBox("Departing City?")
    arriving_date = InputBox("Arriving Date? (ETA)")
    
    
    Dim dow_relative As String
    dow_relative = ""
    
    Dim tomorrows_date As Date
    tomorrows_date = Date
    tomorrows_date = DateAdd("m", 1, tomorrows_date)
    
    If arriving_date = "today" Then
        dow_relative = "today"
        arriving_date = Date
    ElseIf arriving_date = "tomorrow" Then
        dow_relative = "tomorrow"
        arriving_date = tomorrows_date
    ElseIf DateValue(arriving_date) = Date Then
        dow_relative = "today"
    ElseIf DateValue(arriving_date) = tomorrows_date Then
        dow_relative = "tomorrow"
    End If

    
    
    
    Dim subject As String
    subject = "S/ " & company_name & " HAWB# " & po_number
    
    
    Dim message As String
    message = "Good morning, " & personal_name
    message = message & vbNewLine & vbNewLine
    
    message = message & "We wanted to let you know that your airfreight shipment from " & departing_city & " is still due to arrive " & dow_relative & ", " & arriving_date & ". "
    message = message & "We will confirm with the airlines "
    
    If dow_relative = "today" Then
        message = message & "as soon as the flight has departed "
    Else:
        If dow_relative = "tomorrow" Then
            message = message & "tomorrow "
        Else:
            message = message & "as soon as the flight has departed "
        End If
    End If
    
    message = message & "and provide you with additional updates. "
    message = message & vbNewLine & vbNewLine
    
    message = message & "Should you have any questions, please feel free to let us know. "
    message = message & vbNewLine & vbNewLine
    
    message = message & "Thank you and have a great day! "
    message = message & vbNewLine & vbNewLine
    
    

    With OutMail
        .To = ""
        .CC = "siobhansweeney@usffcl.com; kristiponce@usffcl.com"
        .BCC = ""
        .subject = .subject & subject
        .Body = .Body & message
        .Display
        If MsgBox("Send it?", vbYesNo) = vbYes Then .Send
    End With

    Set OutMail = Nothing
    
    
End Sub




