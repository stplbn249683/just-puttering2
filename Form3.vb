Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.IO

Public Class Form3
  Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Dim AppPath$, error1%, sFileName$
    AppPath = Application.StartupPath
    sFileName = AppPath & "\GetStockData.ini"
    error1 = ReadDefaults(sFileName)
    If error1 < 0 Then MessageBox.Show("Error reading file " & sFileName)
  End Sub

  Private Sub butImport_Click(sender As Object, e As EventArgs) Handles butImport.Click
    Dim AppPath$
    AppPath = Application.StartupPath

    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    Application.DoEvents()
    With UserInput
      Call ImportYahooHistorical(.yahoo_historical_input_folder, .yahoo_historical_output_folder)
    End With
    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butUpdate_Click(sender As Object, e As EventArgs) Handles butUpdate.Click
    lblCount.Text = ""
    Me.Cursor = Cursors.WaitCursor
    With UserInput
      Call UpdateDatabase(.yahoo_historical_output_folder, .data_source)
    End With
    Me.Cursor = Cursors.Default
  End Sub
End Class
