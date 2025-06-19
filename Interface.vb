'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Environment
Imports System.Windows.Forms

'This module contains this program's interface.
Public Class InterfaceWindow
   Private ReadOnly TOOL_TIP As New ToolTip()   'Contains this window's tooltip.

   ' This procedure initializes this window.
   Private Sub InterfaceWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
      Try
         With My.Application.Info
            Me.Text = $"{ .Title} v{ .Version} - by: { .CompanyName}"
         End With

         With Screen.PrimaryScreen.Bounds
            Me.Width = CInt(.Width / 1.5)
            Me.Height = CInt(.Height / 1.5)
            Me.Left = CInt((.Width / 2) - (Me.Width / 2))
            Me.Top = CInt((.Height / 2) - (Me.Height / 2))
         End With

         With TOOL_TIP
            .SetToolTip(GUIDListBox, "Enter the GUID's to search for here. Each GUID should be on its own line.")
            .SetToolTip(ResultsBox, "Displays information regarding the specified GUIDs after a search.")
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure adjusts this window to its new size.
   Private Sub InterfaceWindow_Resize(sender As Object, e As EventArgs) Handles Me.Resize
      Try
         GUIDListBox.Width = CInt(Me.ClientSize.Width - 32)
         GUIDListBox.Height = CInt((Me.ClientSize.Height / 2) - 48)

         ResultsBox.Width = CInt(Me.ClientSize.Width - 32)

         ResultsBox.Height = CInt((Me.ClientSize.Height / 2) - 48)
         ResultsBox.Top = CInt(GUIDListBox.Top + GUIDListBox.Height + 48)

         ResultsLabel.Top = CInt(GUIDListBox.Top + GUIDListBox.Height + 16)

         SearchButton.Left = CInt((Me.ClientSize.Width - 32) - SearchButton.Width)
         SearchButton.Top = CInt(GUIDListBox.Top + GUIDListBox.Height + 16)
      Catch
      End Try
   End Sub

   'This procedure gives the command to start searching for the specified GUIDs.
   Private Sub SearchButton_Click(sender As Object, e As EventArgs) Handles SearchButton.Click
      Try
         Dim GUIDText As String = Nothing

         Cursor.Current = Cursors.WaitCursor
         ResultsBox.Text = String.Empty

         For Each Line As String In GUIDListBox.Text.Split({NewLine}, StringSplitOptions.None)
            GUIDText = GetGUIDFromText(Line)
            If String.IsNullOrWhiteSpace(GUIDText) Then GUIDText = Line
            ResultsBox.AppendText(FindGUID(FormatGUID(GUIDText)))
            Application.DoEvents()
         Next Line
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      Finally
         Cursor.Current = Cursors.Default
      End Try
   End Sub
End Class
