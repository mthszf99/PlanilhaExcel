Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim xlWb As Object
        Dim xlApp As Object
        Dim xlWbName As String
        Dim xlWbPath As String

        On Error Resume Next

        xlApp = CreateObject("Excel.Application")

        xlWbName = "Planilha_Segura.xlsb"

        xlWbPath = Application.StartupPath

        Me.Hide()

        xlWb = xlApp.workbooks.open(xlWbPath & "\" & xlWbName)

        xlWb = Nothing
        xlApp = Nothing
        xlWbName = Nothing
        xlWbPath = Nothing

        Application.Exit()





    End Sub
End Class
