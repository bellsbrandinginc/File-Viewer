Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Windows.Threading
Imports System.IO.File
Imports System.Xml
Imports Alphaleonis
Imports System.Text.Encoding
Imports System.Object
Imports System.Text.RegularExpressions
Imports System.Data
Imports System.Windows.Controls.Primitives

Imports System.Linq


Public Class Window1
    Dim myTable As DataTable
    Dim TallyItemsList As List(Of TallyItem)

    Dim tallyTable As DataTable = New DataTable

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.EventArgs)


        collapseExpander.IsExpanded = False

        If (collapseExpander.IsExpanded = False) Then
            dataGrid.MinWidth = "1145"
            dataGrid.MaxWidth = "3250"
            dataGrid.Margin = New Thickness(58, 30, 0, 0)
        End If




        myTable = New DataTable("MyTable")
        Dim i As Integer
        Dim myRow As DataRow
        Dim fieldValues As String()



        Dim inputfile As String = "\\lvdiprodata\EXPORTS01\PS100000\All domains.dat"
        Using fl As FileStream = New FileStream(inputfile, FileMode.Open, FileAccess.Read)
            Using fl_in As StreamReader = New StreamReader(fl, System.Text.Encoding.Unicode)

                Dim line As String = fl_in.ReadLine
                line = line.Replace(Chr(254), "")
                fieldValues = line.Split(Chr(20))
                For i = 0 To fieldValues.Length() - 1
                    myTable.Columns.Add(New DataColumn(fieldValues(i)))


                Next



                While Not (fl_in.Peek() = -1)
                    line = fl_in.ReadLine
                    line = line.Replace(Chr(254), "")
                    fieldValues = line.Split(Chr(20))

                    myRow = myTable.NewRow

                    For i = 0 To fieldValues.Length() - 1
                        myRow.Item(i) = fieldValues(i).ToString
                    Next

                    myTable.Rows.Add(myRow)
                End While

            End Using
        End Using




        dataGrid.ItemsSource = myTable.DefaultView
    End Sub


    Private Sub DataGrid_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs)

        Dim dep As DependencyObject = e.OriginalSource


        While Not dep Is Nothing And Not TypeOf dep Is DataGridCell

            dep = VisualTreeHelper.GetParent(dep)
        End While

        If TypeOf dep Is DataGridCell Then
            Dim gcell As DataGridCell = DirectCast(dep, DataGridCell)

            dataGrid.ContextMenu = New ContextMenu
            Dim mia As MenuItem = New MenuItem
            mia.Header = "Tally " + gcell.Column.Header
            AddHandler mia.Click, AddressOf myEventHandler
            dataGrid.ContextMenu.Items.Add(mia)

            dataGrid.ContextMenu.IsOpen = True


        End If



    End Sub


    Sub myEventHandler(sender As Object, e As System.EventArgs)
        Dim columnheader As String = ""
        If sender.header <> "" Then
            columnheader = sender.header
            columnheader = columnheader.Substring(6)





            Dim query = From row In myTable
                        Group row By columnname = row.Field(Of String)(columnheader).ToUpper Into columnGroup = Group
                        Select New With {
                            Key columnname,
                            .count = columnGroup.Count
                       }

            TallyItemsList = New List(Of TallyItem)

            For Each x In query
                Dim Item As TallyItem = New TallyItem

                Item.ColumnHeader = columnheader
                Item.Name = x.columnname
                Item.Count = x.count


                TallyItemsList.Add(Item)

            Next



            listViewTally.ItemsSource = TallyItemsList

        End If




    End Sub


    Private Sub expanderHasExpanded(ByVal sender As Object, ByVal args As RoutedEventArgs)

        dataGrid.MinWidth = "870"
        If Me.WindowState = System.Windows.WindowState.Maximized Then
            dataGrid.MinWidth = "1575"
            listViewTally.MinHeight = "980"
        End If
        dataGrid.MaxWidth = "3250"
        dataGrid.Margin = New Thickness(348, 30, 0, 0)

    End Sub




    Private Sub expanderHasCollapsed(ByVal sender As Object, ByVal args As RoutedEventArgs)

        dataGrid.MinWidth = "1145"
        dataGrid.MaxWidth = "3250"
        If Me.WindowState = System.Windows.WindowState.Maximized Then
            dataGrid.MinWidth = "1865"
        End If
        dataGrid.Margin = New Thickness(58, 30, 0, 0)
    End Sub

    Protected Overrides Sub OnStateChanged(ByVal e As EventArgs)

        If Me.WindowState = System.Windows.WindowState.Maximized Then
            dataGrid.MinWidth = "1865"
            dataGrid.MaxWidth = "4500"
            dataGrid.MinHeight = "1020"
            dataGrid.Margin = New Thickness(58, 30, 0, 0)

            listViewTally.Margin = New Thickness(526, 15, -590, -384)
            listViewTally.MinHeight = "980"
        End If

        If Me.WindowState = System.Windows.WindowState.Normal Then

            dataGrid.MinWidth = "1145"
            dataGrid.MaxWidth = "3250"
            dataGrid.MinHeight = "741.5"
            dataGrid.Margin = New Thickness(58, 30, 0, 0)
            listViewTally.Margin = New Thickness(168, 20, -232, -389)
            listViewTally.MinHeight = "850"
        End If
    End Sub



    Private Sub Window1_SizeChanged(ByVal sender As Object, ByVal e As SizeChangedEventArgs)

        AddHandler Window1.SizeChanged, AddressOf Window1_SizeChanged
        If Me.Height > 800 Then

            dataGrid.MinWidth = 810
            dataGrid.MaxWidth = 4500

            dataGrid.Margin = New Thickness(58, 30, 0, 0)

        End If

        If Me.Height < 800 Then

            dataGrid.MinWidth = 950

            dataGrid.Margin = New Thickness(0, 0, 0, 0)

        End If



    End Sub
End Class
