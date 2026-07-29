Attribute VB_Name = "global"
Public NeedsHelp As Boolean
Public UserOperation As Integer '1 - deposit, 2 - withdrawal, 3 - transfer, 4 - credit/debit card pin change, 5 - currency exchange
Public isFinished As Boolean
Public isShowingCover As Boolean
Public language As String
Public isAllowed As Boolean
Public globalWidth As Integer
Public globalHeight As Integer

Public Function HasNonNumeric(ByVal InputStr As String)
    HasNonNumeric = Not IsNumeric(CStr(InputStr))
End Function

Public Sub UnloadAllForms()
Unload currencyexchange
Unload usermenu
Unload userinput
Unload changepin
Unload transfer
End Sub

Public Sub LoadAllForms()

    'witdth and height, change before compiling for your screen resolution
    globalWidth = 10240 ' 1024 pixels * 10 twips per pixel
    globalHeight = 7680 ' 768 pixels * 10 twips per pixel
    
    Load currencyexchange
    Load usermenu
    Load userinput
    Load changepin
    Load transfer
    
End Sub


