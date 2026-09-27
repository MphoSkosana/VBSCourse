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


'declare
DIM accNumber
DIM recAccNumber
DIM amount
DIM continueLoop, invalidAcc, invalidRecAcc
DIM strOption
DIM balance


'Loop control variable for options
continueLoop = TRUE

'Loop control variable for account validation
invalidAcc = TRUE


SET objAccount = NEW Account

'get full name 
objAccount.strFullName = InputBox("Please enter your full name:")

'initialise account to 1000
objAccount.dblBalance = 1000

balance = objAccount.dblBalance

'get the owner's account number
DO WHILE invalidAcc

    accNumber = Trim(InputBox("Please enter your account number:"))

    IF Len(accNumber) >= 7 AND Len(accNumber) <= 10 THEN

        objAccount.strAccNum = accNumber

        invalidAcc = FALSE

    ELSE

        MsgBox "Invalid account number! Please try again"

    END IF

LOOP


'keeping track of the balance
MsgBox "Starting Balance: R" & balance


'program starts
DO WHILE continueLoop

    'get option from user
    strOption = InputBox("Please select an option below: " & vbNewLine & _
                        "1 - Withdrawal" & vbNewLine & _
                        "2 - Deposit" & vbNewLine & _
                        "3 - Transfer" & vbNewLine & _
                        "4 - Check Balance" & vbNewLine & _
                        "5 - Exit" )

    'validate options
    IF strOption = "5" THEN

        continueLoop = FALSE

    ELSEIF strOption = "1" OR strOption = "2" OR strOption = "3" OR strOption = "4" THEN

        SELECT CASE(strOption)
        CASE "1"

            DO

                amount = CDbl(InputBox("Enter the amount you wish to withdraw:"))

            LOOP UNTIL amount > 0 AND amount <= balance

            balance = objAccount.withdrawal(amount)

            MsgBox "Withdrawal Complete!" & vbNewLine & "New Balance: R" & balance

        CASE "2"

            DO

                amount = CDbl(InputBox("Enter the amount you wish to deposit:"))

            LOOP UNTIL amount > 0

            balance = objAccount.deposit(amount)

            MsgBox "Deposit Complete!" & vbNewLine & "New Balance: R" & balance

        CASE "3"
            're-initialise control variable
            invalidRecAcc = TRUE

            'validate account number 
            DO WHILE invalidRecAcc

                recAccNumber = Trim(InputBox("Please enter beneficiary account number:"))

                IF Len(recAccNumber) >= 7 AND Len(recAccNumber) <= 10 THEN

                    objAccount.strRecAccNum = recAccNumber

                    invalidRecAcc = FALSE

                ELSE

                    MsgBox "Invalid account number! Please try again"

                END IF

            LOOP

            'get amount
            DO

                amount = CDbl(InputBox("Enter the amount you wish to transfer:"))

            LOOP UNTIL amount > 0 AND amount <= balance

            balance = objAccount.transfer(amount, recAccNumber)

            MsgBox "Transfer Complete!" & vbNewLine & _
                    "Beneficiary Account Number: " & objAccount.getRecAccountNumber() & vbNewLine & _
                    "Amount Transferred: R" & amount & vbNewLine & _
                    "Balance: R" & balance

        CASE "4"

            balance = objAccount.checkBalance()

            MsgBox "Current Balance: R" & balance

        END SELECT

    ELSE

        MsgBox "Invalid option! Please select a valid option."

    END IF

LOOP


MsgBox (objAccount.getAccountDetails())