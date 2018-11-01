Imports System.Windows.Navigation
Imports System.Diagnostics
Imports System.Windows.Controls
Imports System.ComponentModel
Imports System.Windows

Public Class LWFields
    Inherits Window

    Friend closeProgramAlreadyRequested As Boolean = False

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub CheckBox1_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox1 As CheckBox = TryCast(sender, CheckBox)
        parentDateGrid.Visibility = If(checkbox1.IsChecked, Visibility.Visible, Visibility.Collapsed)
        checkBox2.IsChecked = False
        checkBox3.IsChecked = False
    End Sub

    Private Sub CheckBox2_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        checkBox1.IsChecked = False
        Dim checkbox2 As CheckBox = TryCast(sender, CheckBox)
        emailDomainGrid.Visibility = If(checkbox2.IsChecked, Visibility.Visible, Visibility.Collapsed)
        checkBox3.IsChecked = False
    End Sub

    Private Sub CheckBox3_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        checkBox1.IsChecked = False
        checkBox2.IsChecked = False
        Dim checkbox3 As CheckBox = TryCast(sender, CheckBox)
        dateTimeGrid.Visibility = If(checkbox3.IsChecked, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Me.Closed
        closeProgramAlreadyRequested = False
        Me.Close()
    End Sub

    Private Sub pdButton_Click3(sender As Object, e As EventArgs)
        hideParentDateSection.Visibility = System.Windows.Visibility.Visible
        hideParentDateIcon.Visibility = System.Windows.Visibility.Hidden
        showParentDateIcon.Visibility = System.Windows.Visibility.Visible
    End Sub

    Private Sub pdButton_Click3b(sender As Object, e As EventArgs)
        hideParentDateSection.Visibility = System.Windows.Visibility.Hidden
        showParentDateIcon.Visibility = System.Windows.Visibility.Hidden
        hideParentDateIcon.Visibility = System.Windows.Visibility.Visible
    End Sub

    Private Sub edButton_Click3(sender As Object, e As EventArgs)
        hideEmailDomainSection.Visibility = System.Windows.Visibility.Visible
        hideEmailDomainIcon.Visibility = System.Windows.Visibility.Hidden
        showEmailDomainIcon.Visibility = System.Windows.Visibility.Visible
    End Sub

    Private Sub edButton_Click3b(sender As Object, e As EventArgs)
        hideEmailDomainSection.Visibility = System.Windows.Visibility.Hidden
        showEmailDomainIcon.Visibility = System.Windows.Visibility.Hidden
        hideEmailDomainIcon.Visibility = System.Windows.Visibility.Visible
    End Sub

    Private Sub dtButton_Click3(sender As Object, e As EventArgs)
        hideDateTimeSection.Visibility = System.Windows.Visibility.Visible
        hideDateTimeIcon.Visibility = System.Windows.Visibility.Hidden
        showDateTimeIcon.Visibility = System.Windows.Visibility.Visible
    End Sub

    Private Sub dtButton_Click3b(sender As Object, e As EventArgs)
        hideDateTimeSection.Visibility = System.Windows.Visibility.Hidden
        showDateTimeIcon.Visibility = System.Windows.Visibility.Hidden
        hideDateTimeIcon.Visibility = System.Windows.Visibility.Visible
    End Sub

    Private Sub Button_Click4(sender As Object, e As RoutedEventArgs)
        Dim dlg As New Microsoft.Win32.SaveFileDialog()
        dlg.FileName = "Output" ' Default file name
        dlg.DefaultExt = ".csv" ' Default file extension
        dlg.Filter = "Microsoft Excel Comma Separated Values File (.csv)|*.csv" ' Filter files by extension

        ' Show save file dialog box
        Dim result? As Boolean = dlg.ShowDialog()

        ' Process save file dialog box results
        If result = True Then
            ' Save document
            Dim filename As String = dlg.FileName
        End If
    End Sub
End Class

