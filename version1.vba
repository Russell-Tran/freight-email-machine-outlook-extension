Sub MyTemplate4()
    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem ' Reply
    Dim olRecip As Recipient ' Add Recipient



    For Each olItem In Application.ActiveExplorer.Selection
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
    
    olItem.Subject = "PO to " & company_name & " - " & "PO# " & po_number
    
    
    Dim message As String
    message = "Good morning, " & personal_name
    message = message & "<br>" & "<br>"
    
    message = message & "We wanted to let  you know that this PO# " & po_number & " has been booked as follows. "
    message = message & "It currently has an ETA into " & arriving_city & " of " & arriving_date & ". "
    message = message & "Once the documents are available, could you please send them to us? "
    message = message & "<br>" & "<br>"
    
    message = message & "<b>"
    message = message & "VESSEL : " & vessel_name & "<br>"
    message = message & "ETD " & departing_city & " : " & departing_date & "<br>"
    message = message & "ETA " & arriving_city & " : " & arriving_date & "<br>"
    message = message & "Q'ty : " & qty_name & "<br>"
    message = message & "</b>"
    
    Set olReply = olItem.ReplyAll

    olReply.HTMLBody = message & vbCrLf & olReply.HTMLBody
    olReply.Display

    
        'olReply.Send
    Next olItem
End Sub




Sub MyTemplate5()
    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem ' Reply
    Dim olRecip As Recipient ' Add Recipient



    For Each olItem In Application.ActiveExplorer.Selection
    Dim company_name, po_number, personal_name, vessel_number, departing_city, departing_date, arriving_city, arriving_date, qty_name As String
    company_name = InputBox("Company Name?")
    po_number = InputBox("HWAB #?")
    personal_name = InputBox("Personal Name? (First name)")
    departing_city = InputBox("Departing City?")
    arriving_date = InputBox("Arriving Date? (ETA)")
    
    
    Dim dow_relative As String
    dow_relative = ""
    If DateValue(arriving_date) = Date Then dow_relative = "today"
    If DateValue(arriving_date) = Date Then dowa_relative = "tomorrow"
    
    
    
    olItem.Subject = "S/ " & company_name & " HAWB " & "HAWB# " & po_number
    
    
    Dim message As String
    message = "Good morning, " & personal_name
    message = message & "<br>" & "<br>"
    
    message = message & "We wanted to let  you know that your airfreight shipment from " & departing_city & " is still due to arrive " & dow_relative & ", " & arriving_date & ". "
    message = message & "We will confirm with the airlines " & dow_relative & " and provide you with additional updates. "
    message = message & "<br>" & "<br>"
    
    message = message & "Should you have any questions, please feel free to let us know. "
    message = message & "<br>" & "<br>"
    
    message = message & "Thank you and have a great day! "
    message = message & "<br>" & "<br>"
    
    Set olReply = olItem.ReplyAll

    olReply.HTMLBody = message & vbCrLf & olReply.HTMLBody
    olReply.Display

    
        'olReply.Send
    Next olItem
End Sub