Class Account

    public strFullName
    public strAccNum
    public strRecAccNum
    public dblAmnt
    public dblBalance



    public FUNCTION withdrawal(amount)

        dblBalance = dblBalance - amount

        withdrawal = dblBalance

    END FUNCTION



    public FUNCTION deposit(amount)

        dblBalance = dblBalance + amount

        deposit = dblBalance

    END FUNCTION



    public FUNCTION transfer(amount, recAccNum)

        dblBalance = dblBalance - amount

        transfer = dblBalance

    END FUNCTION



    public FUNCTION checkBalance()

        checkBalance = dblBalance

    END FUNCTION



    public FUNCTION getAccountNumber()

        getAccountNumber = "ACC-" & strAccNum

    END FUNCTION



    public FUNCTION getRecAccountNumber()

        getRecAccountNumber = "ACC-" & strRecAccNum

    END FUNCTION



    public FUNCTION getAccountDetails()

    getAccountDetails = "Account Holder: " & strFullName & vbNewLine & _
                        "Account Number: " & getAccountNumber() & vbNewLine & _
                        "Balance: R" & checkBalance()

    END FUNCTION



END Class



'************************************************************************************************************************

'Main Script



Dim objAccount, fullName, accNumber
Dim continueLoop, choice
Dim amount, transferAcc
Dim balance

Set objAccount = New Account

'================================================================================
' Main Menu / Program Start
'================================================================================
continueLoop = True

'Ask user for full name and account number
fullName = InputBox("Enter your full name:")

objAccount.dblBalance = 1000

accNumber = ""

Do
    accNumber = InputBox("Enter your account number (7 to 10 digits):")

    
    If Len(accNumber) >= 7 And Len(accNumber) <= 10 Then
        
        If IsNumeric(accNumber) Then
            Exit Do
        Else
            MsgBox "Invalid account number! Please enter digits only."
        End If
    Else
        MsgBox "Invalid account number! Please try again"
    End If
Loop

objAccount.strFullName = fullName 
objAccount.strAccNum = accNumber

'================================================================================
' Display menu loop
'================================================================================
Do While continueLoop

    choice = InputBox( _
        "Account Application Menu:" & vbNewLine & _
        "1. Withdrawal" & vbNewLine & _
        "2. Deposit" & vbNewLine & _
        "3. Transfer" & vbNewLine & _
        "4. Check balance" & vbNewLine & _
        "5. Exit" & vbNewLine & _
        "Choose an option (1-5):")

    Select Case choice

        Case "1"  'Withdrawal

            amount = 0
            Do
                amount = InputBox("Enter withdrawal amount:")

                If Not IsNumeric(amount) Then
                    MsgBox "Invalid amount! Please enter a numeric value."
                ElseIf CDbl(amount) <= 0 Then
                    MsgBox "Withdrawal amount must be greater than 0."
                ElseIf CDbl(amount) > objAccount.checkBalance() Then
                    MsgBox "Insufficient funds! You cannot withdraw more than your balance."
                Else
                    Exit Do
                End If

            Loop

            
            objAccount.withdrawal CDbl(amount)

            
            MsgBox "Withdrawal Complete!" & vbNewLine & _
                   objAccount.getAccountDetails()

        Case "2"  'Deposit

            amount = 0
            Do
                amount = InputBox("Enter deposit amount:")

                If Not IsNumeric(amount) Then
                    MsgBox "Invalid amount! Please enter a numeric value."
                ElseIf CDbl(amount) <= 0 Then
                    MsgBox "Deposit amount must be greater than 0."
                Else
                    Exit Do
                End If
            Loop

            
            objAccount.deposit CDbl(amount)

            
            MsgBox "Deposit Successful!" & vbNewLine & _
                   objAccount.getAccountDetails()

        Case "3"  'Transfer

            'Receiving account number validation (7-10 digits)
            transferAcc = ""
            Do
                transferAcc = InputBox("Enter the receiving account number (7 to 10 digits):")

                If Len(transferAcc) >= 7 And Len(transferAcc) <= 10 Then
                    If IsNumeric(transferAcc) Then
                        Exit Do
                    Else
                        MsgBox "Invalid receiving account number! Please enter digits only."
                    End If
                Else
                    MsgBox "Invalid account number length! Please try again"
                End If
            Loop

            amount = 0
            Do
                amount = InputBox("Enter transfer amount:")

                If Not IsNumeric(amount) Then
                    MsgBox "Invalid amount! Please enter a numeric value."
                ElseIf CDbl(amount) <= 0 Then
                    MsgBox "Transfer amount must be greater than 0."
                ElseIf CDbl(amount) > objAccount.checkBalance() Then
                    MsgBox "Insufficient funds! You cannot transfer more than your balance."
                Else
                    Exit Do
                End If

            Loop

            'Your original style transfer call (sender subtracts)
            objAccount.transfer CDbl(amount), transferAcc

            
            MsgBox "Transfer Successful!" & vbNewLine & _
                   "Recipient Account Number: ACC-" & transferAcc & vbNewLine & _
                   "Amount Transferred: R" & CDbl(amount) & vbNewLine & _
                   objAccount.getAccountDetails()

        Case "4"  'Check balance

            
            MsgBox "Your Balance:" & vbNewLine & objAccount.getAccountDetails()

        Case "5"  'Exit

            continueLoop = False
            MsgBox "Thank you for using the system!"
            
        Case Else
            MsgBox "Invalid choice! Please choose an option between 1 and 5."

    End Select

Loop


MsgBox (objAccount.getAccountDetails())