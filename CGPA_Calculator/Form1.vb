Public Class Form1

    Dim rows As New Microsoft.VisualBasic.Collection()
    Dim scoreColumn As Integer
    Dim totalUnits As Integer
    Dim totalGradePoint As Integer
    Dim lastRow As Integer
    Dim minimumNumberOfCourses As Integer
    Dim GPA As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        minimumNumberOfCourses = 8

        If (DataGridView1.Rows.Count - 1) < minimumNumberOfCourses Then
            MessageBox.Show("Enter Complete details of at least 8 Courses", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            GoTo endprocess
        End If
        Debug.WriteLine((DataGridView1.Rows.Count - 1))
        lastRow = DataGridView1.Rows.Count - 1
        Try

            For Each currentRow As DataGridViewRow In DataGridView1.Rows

                If currentRow.Index <> lastRow Then

                    totalUnits += currentRow.Cells(3).Value
                    scoreColumn = currentRow.Cells(1).Value
                    Select Case scoreColumn
                        Case 70 To 100
                            currentRow.Cells(2).Value = "A"
                        Case 60 To 69
                            currentRow.Cells(2).Value = "B"
                        Case 50 To 59
                            currentRow.Cells(2).Value = "C"
                        Case 45 To 49
                            currentRow.Cells(2).Value = "D"
                        Case 40 To 44
                            currentRow.Cells(2).Value = "E"
                        Case 0 To 39
                            currentRow.Cells(2).Value = "F"
                        Case Else
                            MessageBox.Show("Score should be between 0 and 100, Please Change!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Select

                    totalGradePoint += Integer.Parse(currentRow.Cells(3).Value) * Integer.Parse(getUnitForGrade(currentRow.Cells(2).Value))

                End If

            Next

            GPA = Double.Parse((totalGradePoint / totalUnits)).ToString("0.00") + " GPA"
            TextBox5.Text = GPA
            totalUnits = 0
            totalGradePoint = 0

        Catch ex As Exception

            MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        MessageBox.Show("Name : " + TextBox1.Text + vbNewLine +
                        vbNewLine + "Faculty : " + TextBox2.Text + vbNewLine +
                        vbNewLine + "Department : " + TextBox3.Text + vbNewLine +
                        vbNewLine + "MatricNo : " + TextBox4.Text + vbNewLine +
                        vbNewLine + "Year : " + DateTimePicker1.Value + vbNewLine +
                        vbNewLine + "GPA : " + GPA, TextBox1.Text, MessageBoxButtons.OK, MessageBoxIcon.None)
endprocess:
    End Sub

    Private Function getUnitForGrade(grade As String)
        Dim point As Integer
        Select Case grade
            Case "A"
                point = 5
            Case "B"
                point = 4
            Case "C"
                point = 3
            Case "D"
                point = 2
            Case "E"
                point = 1
            Case "F"
                point = 0
            Case Else
                Debug.WriteLine("Not between A and F, please change!")
        End Select

        Return point

    End Function


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grade.ReadOnly = True
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
