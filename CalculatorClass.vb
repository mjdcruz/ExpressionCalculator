
Public Class CalculatorClass
    'The highest or lowest values that can be stored by the program
    Dim maxBound As Decimal = 9223372036854775807
    Dim minBound As Decimal = -9223372036854775807

    Dim errorFound As New ErrorCatch

    'Evaluate expression function
    Function evaluateExpression(ByVal inputString As String) As String
        Dim strLen As Integer = Len(inputString)
        'first split string into substrings to obtain operands and numValues
        Dim splitExpression(strLen) As String
        splitExpression = tokenizeInput(inputString, strLen)
        'calculate answer
        Dim answerVal As String = calculateExpression(splitExpression)

        Return answerVal
    End Function

    'splits string into substrings of individual characters to easier sort operands and nums
    Function tokenizeInput(ByVal inputString As String, ByVal stringLen As Integer) As String()

        Dim arrayExpression(stringLen) As String

        For counter As Integer = 0 To stringLen - 1
            arrayExpression(counter) = CStr(inputString.Chars(counter))
        Next

        Return arrayExpression
    End Function

    'driver method for expression set up, precedence, and calculations
    Function calculateExpression(ByVal expression As String()) As String

        Dim calculateable() As String = setUpExpression(expression)
        Dim expressionVal As String = applyPrecedence(calculateable)

        Return expressionVal
    End Function

    'function to seperate nums and operators
    Function setUpExpression(ByVal origExp As String()) As String()
        'convert string array to string list so we can remove items
        Dim origExpLst As List(Of String) = origExp.ToList
        For counter As Integer = 0 To ((origExpLst.Count) - 1)
            'if next element does not exist skip through loop
            If counter >= origExpLst.Count - 1 Then
                Continue For
            End If
            'If number is adjacent to open parenthesis
            If IsNumeric(origExpLst(counter)) And origExpLst(counter + 1) = "(" Or
                IsNumeric(origExpLst(counter)) And origExpLst(counter + 1) = "[" Or
                    IsNumeric(origExpLst(counter)) And origExpLst(counter + 1) = "{" Then
                MsgBox("Must have operator between number and parenthesis/brackets/braces", vbCritical, "Invalid Expression")
                errorFound.restartCalc()
                End
            End If
            'If number has multiple digits or is a decimal
            If IsNumeric(origExpLst(counter)) And IsNumeric(origExpLst(counter + 1)) Then
                origExpLst(counter) = origExpLst(counter) + origExpLst(counter + 1)
                origExpLst.RemoveAt(counter + 1)
                counter = counter - 1
                Continue For
            ElseIf IsNumeric(origExpLst(counter)) And origExpLst(counter + 1) = "." And (counter + 2) < origExpLst.Count Then
                If Not IsNumeric(origExpLst(counter + 2)) Then
                    MsgBox("Invalid character after decimal point", vbCritical, "Invalid Expression")
                    errorFound.restartCalc()
                    End
                End If
                origExpLst(counter) = origExpLst(counter) + origExpLst(counter + 1) + origExpLst(counter + 2)
                origExpLst.RemoveAt(counter + 1)
                origExpLst.RemoveAt(counter + 1)
                counter = counter - 1
                Continue For
            End If
            'If log function or trig function character is found (groups letters into single element in list
            If Char.IsLetter(origExpLst(counter).Chars(0)) And Char.IsLetter(origExpLst(counter + 1)) Then
                origExpLst(counter) = origExpLst(counter) + origExpLst(counter + 1)
                origExpLst.RemoveAt(counter + 1)
                counter = counter - 1
                Continue For
            End If
        Next
        Return origExpLst.ToArray
    End Function

    'function to execute functions by precedence
    Function applyPrecedence(ByVal groupedExpression As String()) As String

        Dim expression As List(Of String) = groupedExpression.ToList
        'counter to end while if it runs into errors or an infinite loop
        Dim errorCounter As Integer = 0
        Dim openBrackets As New List(Of String)({"(", "[", "{"})
        Dim closeBrackets As New List(Of String)({")", "]", "}"})
        Dim functions() As String = {"sin", "cos", "tan", "cot", "arcsin", "arccos", "arctan", "arcctg", "ln", "log"}
        Dim highOperands() As String = {"^"}
        Dim midOperands() As String = {"*", "/"}
        Dim lowOperands() As String = {"+", "-"}

        expression.RemoveAt(expression.Count - 1)
        'while expression contains more than one element
        While expression.Count > 1
            'PARENTHESIS/BRACKETS/BRACES
            'If the element between parenthesis/brackets/braces is just a number, eliminate parenthesis
            For counter As Integer = 0 To ((expression.Count) - 1)
                If counter > expression.Count - 1 Then
                    Continue For
                End If
                If openBrackets.Contains(expression(counter)) And (counter + 2) < expression.Count Then
                    If closeBrackets.Contains(expression(counter + 2)) Then
                        expression.RemoveAt(counter)
                        expression.RemoveAt(counter + 1)
                    End If
                End If
            Next

            'NEGATE, we must apply negations first to distinguish when "-" means negate numbers and subtract
            For counter As Integer = 0 To ((expression.Count) - 1)
                If counter >= expression.Count - 1 Then
                    Continue For
                End If
                'if negative sign is at the start of an expression
                If counter = 0 And expression(counter).Contains("-") And IsNumeric(expression(counter + 1)) Then
                    expression(counter) = negate(expression(counter + 1))
                    expression.RemoveAt(counter + 1)
                    'if "-" comes after an operand or open bracket
                ElseIf highOperands.Contains(expression(counter)) Or midOperands.Contains(expression(counter)) Or
                    lowOperands.Contains(expression(counter)) Or openBrackets.Contains(expression(counter)) And (counter + 2) < expression.Count Then
                    If expression(counter + 1).Contains("-") Then
                        If IsNumeric(expression(counter + 2)) Then
                            expression(counter + 1) = expression(counter + 1) + expression(counter + 2)
                            expression.RemoveAt(counter + 2)
                        End If
                    End If
                End If
            Next
            'FUNCTIONS, after negations if trig function is found and next element is numeric, performOperation
            For counter As Integer = 0 To ((expression.Count) - 1)
                If counter >= expression.Count - 1 Then
                    Continue For
                End If
                If functions.Contains(expression(counter)) And IsNumeric(expression(counter + 1)) Then
                    expression(counter) = performFunction(expression(counter), expression(counter + 1))
                    expression.RemoveAt(counter + 1)
                End If
            Next
            'EXPONENTS (highOperands)
            For counter As Integer = 0 To ((expression.Count) - 1)
                If counter >= expression.Count - 1 Then
                    Continue For
                End If
                If IsNumeric(expression(counter)) And highOperands.Contains(expression(counter + 1)) And (counter + 2) < expression.Count Then
                    If IsNumeric(expression(counter + 2)) Then
                        expression(counter) = performOperation(expression(counter), expression(counter + 1), expression(counter + 2))
                        expression.RemoveAt(counter + 1)
                        expression.RemoveAt(counter + 1)
                    End If
                End If
            Next
            'MULTIPLY/DIVIDE (midOperands)
            For counter As Integer = 0 To ((expression.Count) - 1)
                If counter >= expression.Count - 1 Then
                    Continue For
                End If
                If IsNumeric(expression(counter)) And midOperands.Contains(expression(counter + 1)) And (counter + 2) < expression.Count Then
                    If IsNumeric(expression(counter + 2)) Then
                        expression(counter) = performOperation(expression(counter), expression(counter + 1), expression(counter + 2))
                        expression.RemoveAt(counter + 1)
                        expression.RemoveAt(counter + 1)
                    End If
                End If
            Next
            'ADD/SUBTRACT (lowOperands)
            For counter As Integer = 0 To ((expression.Count) - 1)
                If counter >= expression.Count - 1 Then
                    Continue For
                End If
                If IsNumeric(expression(counter)) And lowOperands.Contains(expression(counter + 1)) And (counter + 2) < expression.Count Then
                    If IsNumeric(expression(counter + 2)) Then
                        expression(counter) = performOperation(expression(counter), expression(counter + 1), expression(counter + 2))
                        expression.RemoveAt(counter + 1)
                        expression.RemoveAt(counter + 1)
                    End If
                End If
            Next

            errorCounter = errorCounter + 1
            'if runs if calculation results in infinite loop or is unable to computer anything
            If errorCounter >= 1000 Then
                MsgBox("Expression is too long or invalid expression syntax enountered", vbCritical, "Invalid Expression")
                errorFound.restartCalc()
                End
            End If
        End While

        Return String.Join(", ", expression.ToArray())
    End Function

    'function to apply operators
    Function performOperation(ByVal num1 As String, ByVal operand As String, ByVal num2 As String) As String
        Dim value As Double
        'Check if num1 or num2 is too big to be stored
        Try
            inBounds(num1)
        Catch ex As Exception
            MsgBox("Number is too large to be stored!", vbCritical, "Invalid Expression")
            errorFound.restartCalc()
            End
        End Try
        Try
            inBounds(num2)
        Catch ex As Exception
            MsgBox("Number is too large to be stored!", vbCritical, "Invalid Expression")
            errorFound.restartCalc()
            End
        End Try

        'operation statements
        Select Case operand
            Case "^"
                value = Math.Pow(CDec(num1), CDec(num2))
            Case "*"
                value = CDec(num1) * CDec(num2)
            Case "/"
                value = CDec(num1) / CDec(num2)
            Case "+"
                value = CDec(num1) + CDec(num2)
            Case "-"
                value = CDec(num1) - CDec(num2)
            Case Else
                MsgBox("Invalid Expression: Invalid operand detected")
                errorFound.restartCalc()
                End
        End Select
        Return (CStr(value))
    End Function

    Function performFunction(ByVal func As String, ByVal num As String) As String
        Dim value As Double
        'Check if num is too big to be stored
        Try
            inBounds(num)
        Catch ex As Exception
            MsgBox("Number is too large to be stored!", vbCritical, "Invalid Expression")
            errorFound.restartCalc()
            End
        End Try
        'Calcultor always reads num as radians and returns radians when using trig functions

        Select Case func
            Case "sin"
                value = Math.Sin(CDec(num))
            Case "cos"
                value = Math.Cos(CDec(num))
            Case "tan"
                value = Math.Tan(CDec(num))
            Case "arcsin"
                value = Math.Asin(CDec(num))
            Case "arccos"
                value = Math.Acos(CDec(num))
            Case "arctan"
                value = Math.Atan(CDec(num))
            Case "cot"
                value = (Math.Cos(CDec(num))) / (Math.Sin(CDec(num)))
            Case "arcctg"
                value = (Math.PI / 2) - Math.Atan(CDec(num))
            Case "ln"
                value = Math.Log(CDec(num))
            Case "log"
                value = Math.Log10(CDec(num))
            Case Else
                MsgBox("Invalid Expression: Invalid function detected")
                errorFound.restartCalc()
                End
        End Select
        Return (CStr(value))
    End Function

    'negate function
    Function negate(ByVal num As String) As String
        Dim value As Decimal
        value = num - (num * 2)
        Return CStr(value)
    End Function

    Function inBounds(ByVal num As Decimal) As Boolean
        If num > maxBound Or num < minBound Then
            Return False
        End If
        Return True
    End Function
End Class
