Public Class ErrorCatch

    Sub restartCalc()
        Dim answer As Integer

        answer = MsgBox("Would you like to try another expression?", vbQuestion + vbYesNo + vbDefaultButton2, "Restart Calculator")

        If answer = vbNo Then
            End
        End If

        Application.Restart()

    End Sub

End Class
