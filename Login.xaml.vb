Imports System.Diagnostics
Imports System.Windows.Navigation
Imports System.Windows.Controls
Imports System.ComponentModel
Imports System.Windows

Public Class Login
    Inherits Window

    Friend closeProgramAlreadyRequested As Boolean = False

    Public Sub New()
        ' navService = NavigationService.GetNavigationService(Me)
        InitializeComponent()
    End Sub

    Private Sub Button_Click1(sender As Object, e As EventArgs) Handles Me.Closed
        closeProgramAlreadyRequested = False
        Me.Close()
    End Sub

    Private Sub LoginButton(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim mw As MainWindow = New MainWindow()
        mw.Show()
    End Sub
End Class