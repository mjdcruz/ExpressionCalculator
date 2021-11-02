'Mathew Dela Cruz
'Expression Calculator

Public Class ExpressionCalculatorApp

    Dim parenthesisCount As Integer = 0
    Dim bracketCount As Integer = 0
    Dim bracesCount As Integer = 0
    Dim result As String

    'Method for Number and Operator button clicks
    Private Sub BtnClickMethod(sender As Object, e As EventArgs) Handles Num1Btn.Click, Num2Btn.Click,
        Num3Btn.Click, Num4Btn.Click, Num5Btn.Click, Num6Btn.Click, Num7Btn.Click, Num8Btn.Click,
        Num9Btn.Click, Num0Btn.Click, PlusBtn.Click, MinusBtn.Click, MultiplyBtn.Click,
        DivideBtn.Click, ExpBtn.Click, DecimalBtn.Click

        Dim button As Button = CType(sender, Button)

        If button.Name = "Num1Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "1"
        End If
        If button.Name = "Num2Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "2"
        End If
        If button.Name = "Num3Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "3"
        End If
        If button.Name = "Num4Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "4"
        End If
        If button.Name = "Num5Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "5"
        End If
        If button.Name = "Num6Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "6"
        End If
        If button.Name = "Num7Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "7"
        End If
        If button.Name = "Num8Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "8"
        End If
        If button.Name = "Num9Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "9"
        End If
        If button.Name = "Num0Btn" Then
            InputTxtBox.Text = InputTxtBox.Text + "0"
        End If
        If button.Name = "PlusBtn" Then
            InputTxtBox.Text = InputTxtBox.Text + "+"
        End If
        If button.Name = "MinusBtn" Then
            InputTxtBox.Text = InputTxtBox.Text + "-"
        End If
        If button.Name = "MultiplyBtn" Then
            InputTxtBox.Text = InputTxtBox.Text + "*"
        End If
        If button.Name = "DivideBtn" Then
            InputTxtBox.Text = InputTxtBox.Text + "/"
        End If
        If button.Name = "DecimalBtn" Then
            InputTxtBox.Text = InputTxtBox.Text + "."
        End If
        If button.Name = "ExpBtn" Then
            InputTxtBox.Text = InputTxtBox.Text + "^"
        End If
    End Sub

    'Method for ClrBtn that clears InputTxtBox field
    Private Sub ClrBtn_Click(sender As Object, e As EventArgs) Handles ClrBtn.Click
        InputTxtBox.Text = ""
        parenthesisCount = 0
        bracketCount = 0
        bracesCount = 0
    End Sub

    'Method for trigonemtric functions, log functions and brackets clicks
    Private Sub HighPrecedenceBtn_Click(sender As Object, e As EventArgs) Handles ParenthesisBtn.Click,
        BracketBtn.Click, BracesBtn.Click, SinBtn.Click, CosBtn.Click, TanBtn.Click, ASinBtn.Click,
        ACosBtn.Click, ATanBtn.Click, CotBtn.Click, ACotTanBtn.Click, NLogBtn.Click, LogBtn.Click

        Dim button As Button = CType(sender, Button)
        'Parenthesis, Brackets, Braces
        If button.Name = "ParenthesisBtn" Then
            If parenthesisCount > 0 Then
                parenthesisCount = parenthesisCount - 1
                InputTxtBox.Text = InputTxtBox.Text + ")"
            Else
                parenthesisCount = parenthesisCount + 1
                InputTxtBox.Text = InputTxtBox.Text + "("
            End If
        End If
        If button.Name = "BracketBtn" Then
            If bracketCount > 0 Then
                bracketCount = bracketCount - 1
                InputTxtBox.Text = InputTxtBox.Text + "]"
            Else
                bracketCount = bracketCount + 1
                InputTxtBox.Text = InputTxtBox.Text + "["
            End If
        End If
        If button.Name = "BracesBtn" Then
            If bracesCount > 0 Then
                bracesCount = bracesCount - 1
                InputTxtBox.Text = InputTxtBox.Text + "}"
            Else
                bracesCount = bracesCount + 1
                InputTxtBox.Text = InputTxtBox.Text + "{"
            End If
        End If

        'TrigonemtricFunctions and LogFunctions
        If button.Name = "SinBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "sin("
        End If
        If button.Name = "CosBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "cos("
        End If
        If button.Name = "TanBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "tan("
        End If
        If button.Name = "ASinBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "arcsin("
        End If
        If button.Name = "ACosBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "arccos("
        End If
        If button.Name = "ATanBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "arctan("
        End If
        If button.Name = "CotBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "cot("
        End If
        If button.Name = "ACotTanBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "arcctg("
        End If
        If button.Name = "NLogBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "ln("
        End If
        If button.Name = "LogBtn" Then
            parenthesisCount = parenthesisCount + 1
            InputTxtBox.Text = InputTxtBox.Text + "log("
        End If
    End Sub

    'Calculate expression 
    Private Sub EqualsBtn_Click(sender As Object, e As EventArgs) Handles EqualsBtn.Click

        Dim calculator As New CalculatorClass
        result = calculator.evaluateExpression(InputTxtBox.Text)
        parenthesisCount = 0
        bracketCount = 0
        bracesCount = 0
        InputTxtBox.Text = result

    End Sub
End Class
