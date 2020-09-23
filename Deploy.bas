Attribute VB_Name = "Deploy"
'############################################
'############################################
'###                                      ###
'###   ANGELS DINER CHECKOUT AND TABLES   ###
'###                                      ###
'############################################
'############################################

'############################################
'#                                          #
'#     Written By MICHAEL HOPKINS           #
'#                                          #
'############################################
'#                                          #
'#     Deploy(Module):                      #
'#                                          #
'############################################

'All arrays are made public to the whole program
'usrStore is the Username array
'pswStore is the Password array
'menuStore is the Menu array
'priceStore is the Price array

Public usrStore(20) As String, pswStore(20) As String
Public menuStore(40) As String, priceStore(40) As String

Sub printer1()

'Sub actually prints a hard-copy receipt for
'the customer.

Dim intCtr As Integer
Dim intListCount As Integer

intCtr = 0
intListCount = frmMain.lstFood.ListCount

'Sets font properties
Printer.Font = tahoma
Printer.FontSize = 9
Printer.FontBold = True
Printer.Print "ANGELS DINER"
Printer.FontBold = False
Printer.Print "Customer Copy"
Printer.Print "Keep for your records"

'Loops through purchases printing each a
'line at a time
For intCtr = 0 To (intListCount - 1)
    Printer.Print frmMain.lstFood.List(intCtr)
Next intCtr

Printer.Print "Total Price"
Printer.Print txtTotal.Text
Printer.Print ""
Printer.Print "Change Due"
Printer.Print txtValue.Text
Printer.Print ""
Printer.Print "Thank-you, please"
Printer.Print "call in again."

End Sub

Sub cancelAll()

'clears all double-precision datatype locations in calculator
frmMain.txtValue.Text = "0.00"
store = "0"

End Sub

Sub totPrice()

'############################
'#                          #
'# Calculates Total Price   #
'#                          #
'############################
'#                          #
'# Searches through every   #
'# char of every item in    #
'# the list, searching for  #
'# prices and keeping a     #
'# running total            #
'#                          #
'############################
    
'Searches through each character in a string
Dim strListCtr As Integer
'For each item in the list
Dim intCtr As Integer
'Length of the list
Dim lstLength As Integer
'Length of active item in list
Dim strLength As Integer

'Total price
Dim dblNewPrice As Double

'Sets initial states of above
intCtr = 0
strListCtr = 0
lstLength = frmMain.lstFood.ListCount - 1
strLength = Len(frmMain.lstFood.List(strListCtr))

dblNewPrice = 0

'Main section of sub searches through every character
'of every item in the list searching for a "-" when this is found
'the string has the following string manipulations performed:

'LTrim - trims string of empty spaces
'Mid - trims the current item in the list
'starting at current strListCtr value + 1
'until the end of the string

'########################################
'#                                      #
'#  NOTE: there is no conversion        #
'#  as VB can perform it on-the-fly     #
'#                                      #
'########################################

'This returns the price of the item on the list
'This price is then added to the current running total
For intCtr = 0 To lstLength
    For strListCtr = 1 To strLength
        If Mid(frmMain.lstFood.List(intCtr), strListCtr, 1) = "-" Then
            dblNewPrice = dblNewPrice + Val(LTrim(Mid(frmMain.lstFood.List(intCtr), (strListCtr + 1), (strLength - strListCtr))))
            Exit For
        End If
    Next strListCtr
Next intCtr

'Price is outputted to txtTotal
frmMain.txtTotal.Text = dblNewPrice

End Sub

Sub userDefined()

'Sets state of calculator buttons
With frmMain
        .cmdAddition.Enabled = False
        .cmdSubtract.Enabled = False
        .cmdMultiply.Enabled = False
        .cmdDivide.Enabled = False
        .cmdSendReceipt.Enabled = False
        .cmdDecimal.Enabled = False
        .cmdClear.Enabled = False
        .cmdEquals.Enabled = False
        .cmdCancelUser.Enabled = False
        .cmdUserExit.Enabled = False
        .cmdRoot.Enabled = False
        .cmdPercent.Enabled = False
End With
    
Dim intCtr As Integer
intCtr = 0
For intCtr = 0 To 9
    frmMain.cmdInteger(intCtr).Enabled = False
    Next intCtr
Exit Sub

End Sub

Sub menuLoad()

'Loads the menu and price from file into the arrays
Dim intCtr As Integer
Let intCtr = 0

'Opens the text file containing the menu so its
'data can be inputted into array
Open "d:\diner\menu.txt" For Input As #2

'Loops until the file finishes, entering the data
'into two seperate one-dimensional arrays
Do Until (EOF(2) = True)
    Input #2, menuStore(intCtr), priceStore(intCtr)
    intCtr = intCtr + 1
Loop

'closes text file
Close #2

End Sub


Sub usrLoad()

'Loads the username and password from file into the arrays
Dim intCtr As Integer
Let intCtr = 0

'Opens the text file containing the users so its
'data can be inputted into array
Open "d:\diner\users.txt" For Input As #1

'Loops until the file finishes, entering the data
'into two seperate one-dimensional arrays
Do Until (EOF(1) = True)
    Input #1, usrStore(intCtr), pswStore(intCtr)
    intCtr = intCtr + 1
Loop

'closes text file
Close #1

End Sub

Sub quit()

'unloads all project forms from memory
Unload frmLogin
Unload frmMain

End Sub

Sub addUsers()

'strNewUser - new username
'strNewPsw - new users password
'strNewPswConf - confirms new users password
Dim strNewUser As String, strNewPsw As String, strNewPswConf As String

'Opens user file for to add entry
Open "d:\diner\users.txt" For Append As #1

'Enter new username
strNewUser = InputBox("Enter New Username", "Username")
If strNewUser = "" Then
    MsgBox "No New Users Were Added", vbInformation, "Action Cancelled"
    Close #1
    Exit Sub
End If

'Ask user to enter new password twice
strNewPsw = InputBox("Please Enter Password", "Password")
strNewPswConf = InputBox("Please Confirm the Password", "Password")

'Loop until the two passwords match, if they
'match already then loop is by-passed
Do Until (strNewPsw = strNewPswConf)
    MsgBox "Passwords did not match", vbExclamation, "Please re-enter"
    strNewPsw = InputBox("Please Enter Password Again", "Password")
    strNewPswConf = InputBox("Please Confirm the Password", "Password")
Loop

'Writes Username and Password into file in uppercase
Write #1, UCase(strNewUser), UCase(strNewPsw)

'Closes file
Close #1

'Message box confirms new user added
MsgBox "New user " & UCase(strNewUser) & " added", , "Complete"

End Sub
