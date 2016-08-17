
Imports System.Drawing.Printing.PrinterSettings
Imports System.Management

Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Printers As New Printing.PrintDocument()
        Dim oPS As New System.Drawing.Printing.PrinterSettings
        Dim DefaultPrinterName As String
        Dim PrinterName = Printers.PrinterSettings.PrinterName

        For Each PrinterName In InstalledPrinters()
            ListBox1.Items.Add(PrinterName)
        Next
        Try
            ListBox1.SelectedItem = oPS.PrinterName
        Catch ex As System.Exception
            DefaultPrinterName = ""
        Finally
            oPS = Nothing
        End Try
    End Sub

    Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Boolean
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Printers As New Printing.PrintDocument()
        Dim PrinterName = ListBox1.SelectedItem
        SetDefaultPrinter(PrinterName)
        Close()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class


