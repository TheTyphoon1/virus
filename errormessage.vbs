    Dim windowTitle, messageText
    windowTitle = InputBox("Enter the window title for the message box:")
    messageText = InputBox("Enter the message text for the message box:")
Do    
    MsgBox messageText, vbInformation, windowTitle
Loop