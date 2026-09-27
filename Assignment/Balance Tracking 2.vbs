'************************************************************************************************************************
' Class Account (Updated)
'************************************************************************************************************************
Class Account
    Public strFullName
    Public strAccNum
    Public dblAmnt
    Public dblBalance
    Public strRecAccNum

    '---------------------------
    ' Core balance operations
    '---------------------------

    Public Function withdrawal(amount)
        dblBalance = dblBalance - amount
        withdrawal = dblBalance
    End Function

    Public Function deposit(amount)
        dblBalance = dblBalance + amount
        deposit = dblBalance
    End Function

    Public Function transfer(amount, recAccNum)
        
        dblBalance = dblBalance - amount
        strRecAccNum = recAccNum
        transfer = dblBalance
    End Function

    '---------------------------
    ' Get/check functions
    '---------------------------

    Public Function checkBalance()
        checkBalance = dblBalance
    End Function

    Public Function getAccountNumber()
        getAccountNumber = "ACC-" & strAccNum
    End Function

    Public Function getRecAccountNumber()
        getRecAccountNumber = "ACC-" & strRecAccNum
    End Function

    Public Function getAccountDetails()
        getAccountDetails = "Account Holder: " & strFullName & vbNewLine & _
                            "Account Number: " & getAccountNumber() & vbNewLine & _
                            "Balance: R" & checkBalance()
    End Function

    '************************************************************************************************************************
    ' Input / Validation Functions (Moved from Select Case loops)
    '************************************************************************************************************************

    Public Function PromptAndValidateAccountNumber()
        Dim accNumber
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

        PromptAndValidateAccountNumber = accNumber
    End Function

    Public Function PromptAndValidateMenuChoice()
        Dim choice
        choice = ""

        Do
            choice = InputBox( _
                "Account Application Menu:" & vbNewLine & _
                "1. Withdrawal" & vbNewLine & _
                "2. Deposit" & vbNewLine & _
                "3. Transfer" & vbNewLine & _
                "4. Check balance" & vbNewLine & _
                "5. Exit" & vbNewLine & _
                "Choose an option (1-5):")

            If choice = "1" Or choice = "2" Or choice = "3" Or choice = "4" Or choice = "5" Then
                Exit Do
            Else
                MsgBox "Invalid choice! Please choose an option between 1 and 5."
            End If
        Loop

        PromptAndValidateMenuChoice = choice
    End Function

    Public Function PromptAndValidateWithdrawalAmount()
        Dim amount
        amount = 0

        Do
            amount = InputBox("Enter withdrawal amount:")

            If Not IsNumeric(amount) Then
                MsgBox "Invalid amount! Please enter a numeric value."
            ElseIf CDbl(amount) <= 0 Then
                MsgBox "Withdrawal amount must be greater than 0."
            ElseIf CDbl(amount) > checkBalance() Then
                MsgBox "Insufficient funds! You cannot withdraw more than your balance."
            Else
                Exit Do
            End If
        Loop

        PromptAndValidateWithdrawalAmount = CDbl(amount)
    End Function

    Public Function PromptAndValidateDepositAmount()
        Dim amount
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

        PromptAndValidateDepositAmount = CDbl(amount)
    End Function

    Public Function PromptAndValidateTransferReceivingAccount()
        Dim transferAcc
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

        PromptAndValidateTransferReceivingAccount = transferAcc
    End Function

    Public Function PromptAndValidateTransferAmount()
        Dim amount
        amount = 0

        Do
            amount = InputBox("Enter transfer amount:")

            If Not IsNumeric(amount) Then
                MsgBox "Invalid amount! Please enter a numeric value."
            ElseIf CDbl(amount) <= 0 Then
                MsgBox "Transfer amount must be greater than 0."
            ElseIf CDbl(amount) > checkBalance() Then
                MsgBox "Insufficient funds! You cannot transfer more than your balance."
            Else
                Exit Do
            End If
        Loop

        PromptAndValidateTransferAmount = CDbl(amount)
    End Function

    '************************************************************************************************************************
    ' Action functions (called after validation)
    '************************************************************************************************************************

    Public Function PerformWithdrawal(amount)
        PerformWithdrawal = withdrawal(CDbl(amount))
    End Function

    Public Function PerformDeposit(amount)
        PerformDeposit = deposit(CDbl(amount))
    End Function

    Public Function PerformTransfer(amount, recAccNum)
        PerformTransfer = transfer(CDbl(amount), recAccNum)
    End Function

End Class

'************************************************************************************************************************
' Main Script (Updated)
'************************************************************************************************************************
Dim objAccount, fullName
Dim continueLoop, choice
Dim amount, transferAcc

Set objAccount = New Account

' Ask user for full name
fullName = InputBox("Enter your full name:")

' set initial balance
objAccount.dblBalance = 1000
objAccount.strFullName = fullName

' validate account number using class method
objAccount.strAccNum = objAccount.PromptAndValidateAccountNumber()

continueLoop = True

Do While continueLoop
    choice = objAccount.PromptAndValidateMenuChoice()

    Select Case choice

        Case "1" ' Withdrawal
            amount = objAccount.PromptAndValidateWithdrawalAmount()
            objAccount.PerformWithdrawal(amount)

            MsgBox "Withdrawal Complete!" & vbNewLine & _
                   objAccount.getAccountDetails()

        Case "2" ' Deposit
            amount = objAccount.PromptAndValidateDepositAmount()
            objAccount.PerformDeposit(amount)

            MsgBox "Deposit Successful!" & vbNewLine & _
                   objAccount.getAccountDetails()

        Case "3" ' Transfer
            transferAcc = objAccount.PromptAndValidateTransferReceivingAccount()
            amount = objAccount.PromptAndValidateTransferAmount()

            objAccount.PerformTransfer amount, transferAcc

            MsgBox "Transfer Successful!" & vbNewLine & _
                   "Recipient Account Number: ACC-" & transferAcc & vbNewLine & _
                   "Amount Transferred: R" & CDbl(amount) & vbNewLine & _
                   objAccount.getAccountDetails()

        Case "4" ' Check balance
            MsgBox "Your Balance:" & vbNewLine & objAccount.getAccountDetails()

        Case "5" ' Exit
            continueLoop = False
            MsgBox "Thank you for using the system!"

    End Select
Loop

MsgBox objAccount.getAccountDetails()