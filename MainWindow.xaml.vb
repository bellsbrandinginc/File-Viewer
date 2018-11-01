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



Public Class MainWindow
    Inherits MetroWindow

    Dim field As New Object
    Dim rowIndex As Integer = 0
    Dim _globaltotallinecount As Integer = 0
    Dim _globalerror As Integer = 0
    Dim _currentItemCount As Integer = 0
    Dim _processerror As Integer = 0
    Dim _inputFileList As String
    Dim _isHeader As Boolean
    Dim _BegDoc As String
    Dim _BegAttach As String
    Dim _analyzeMethod As Integer = 0
    Dim _OutFolder As String
    Public Tag2FieldDocCount As Integer
    Dim mainXMLFile As String
    Dim splitFileNamePart As Integer
    Dim WithEvents BackgroundWorker1 As BackgroundWorker
    Dim mainHeaderline As String = ""
    Dim pdButtonClickCount As Integer = 0
    Public LoadFilesItemList As New List(Of LoadFileItem)
    Public Property FileItemsList As New ObservableCollection(Of FileItem)
    Friend closeProgramAlreadyRequested As Boolean = False
    Public userID As String
    Public userPassword As String

    Dim _nativeFileTypeList As ArrayList = New ArrayList
    Dim _knownInternalDomains As ArrayList = New ArrayList




    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TB_Output.Text = ""
        rb_seperate.IsChecked = True
        rb_thirdparty.IsChecked = True
        Me.Height = "400"
        edFileType.Text = "Please select NativeFileType"

        If (Me.AnalyzerTab.IsSelected) Then
            Me.Height = "400"
        End If

        If (Me.CalculateTab.IsSelected) Then
            Me.Height = "800"
        End If

        _nativeFileTypeList.Clear()
        _knownInternalDomains.Clear()

        _nativeFileTypeList.Add("1143")
        _nativeFileTypeList.Add("1602")
        _nativeFileTypeList.Add("1196")
        _nativeFileTypeList.Add("1141")

        _knownInternalDomains.Add("contractor.lw.com".ToLower)
        _knownInternalDomains.Add("dms.asia.lw.com".ToLower)
        _knownInternalDomains.Add("dms.eu.lw.com".ToLower)
        _knownInternalDomains.Add("dms.me.lw.com".ToLower)
        _knownInternalDomains.Add("dms.us.lw.com".ToLower)
        _knownInternalDomains.Add("exchange.lw.com".ToLower)
        _knownInternalDomains.Add("lathamalumni.com".ToLower)
        _knownInternalDomains.Add("lathamcares.org".ToLower)
        _knownInternalDomains.Add("lvdcuc.lw.com".ToLower)
        _knownInternalDomains.Add("lw.com".ToLower)
        _knownInternalDomains.Add("lw.communicate.com".ToLower)
        _knownInternalDomains.Add("lwlocal.com".ToLower)
        _knownInternalDomains.Add("moss.lw.com".ToLower)
        _knownInternalDomains.Add("printeron.lw.com".ToLower)
        _knownInternalDomains.Add("retiredpartner.lw.com".ToLower)
        _knownInternalDomains.Add("vm.asia.lw.com".ToLower)
        _knownInternalDomains.Add("vm.eu.lw.com".ToLower)
        _knownInternalDomains.Add("vm.me.lw.com".ToLower)
        _knownInternalDomains.Add("vm.na.lw.com".ToLower)
        _knownInternalDomains.Add("service-Now.com".ToLower)
        _knownInternalDomains.Add("chromefile.com".ToLower)

    End Sub

    Private Sub MyTabControl_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabControl.MouseEnter
        If (Me.AnalyzerTab.IsSelected) Then
            Me.Height = "400"
            Me.CalculateTab.IsEnabled = False
        End If

        If (Me.CalculateTab.IsSelected) Then
            Me.Height = "800"
        End If
    End Sub

    Private Async Sub BT_Go_Click(sender As Object, e As RoutedEventArgs) Handles BT_Go.Click
        GridStatus.Visibility = System.Windows.Visibility.Visible
        Cursor = Cursors.Wait
        _globaltotallinecount = 0
        _globalerror = 0

        If rb_group.IsChecked = True Then
            _analyzeMethod = 1
        Else
            _analyzeMethod = 0
        End If

        LoadFilesItemList.Clear()
        _processerror = 0
        splitFileNamePart = 1
        progressbar1.Value = 0
        'progressbar1.Maximum = 4

        If TB_Output.Text = "" Then
            Await ShowMessageAsync("", "Please Select report To save.")
            Exit Sub
        ElseIf FileItemsList Is Nothing Or FileItemsList.Count < 1 Then
            Await ShowMessageAsync("", "Please add load file(s) To analyze.")
            Exit Sub
        End If

        For Each item In FileItemsList
            If IsNothing(item.FileEncodingType.SelectedItem) Then
                Await ShowMessageAsync("", "Please make sure encoding Is selected For all load files To be analyzed.")
                Exit Sub
            End If
        Next

        For Each item In FileItemsList
            Dim f As New LoadFileItem
            f.FileName = item.FileName.Text
            f.FileEncodingType = item.FileEncodingType.SelectedItem.ToString
            LoadFilesItemList.Add(f)
        Next

        'mainXMLFile = TB_XML.Text

        _OutFolder = TB_Output.Text
        windowgrip.IsEnabled = False
        Me.BackgroundWorker1 = New BackgroundWorker
        GridStatus.Visibility = System.Windows.Visibility.Visible
        'Enable progress reporting
        BackgroundWorker1.WorkerReportsProgress = True
        'Start the work 
        BackgroundWorker1.RunWorkerAsync()
        'DoWork Event occurs
        'Now control will goes to worker_DoWork Sub because it handles the DoWork Event


    End Sub

    Private Sub worker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        If _analyzeMethod = 0 Then
            ReadLoadFileIndividually()
        Else
            ReadLoadFileGroup()
        End If
    End Sub

    Function ReadLoadFileGroup() As Integer
        Dim outdir As String = Path.GetDirectoryName(_OutFolder)
        Dim filecount As Integer = 1
        Dim outputfile As String = _OutFolder
        Dim MetaFieldItemsList As New List(Of MetaFieldItem)
        Dim MetadataFieldCount As Integer = 0


        Using fout As FileStream = New FileStream(outdir + "\_Errors.Log", FileMode.Create, FileAccess.Write)
            Using fstr_out As StreamWriter = New StreamWriter(fout, System.Text.Encoding.UTF8)

                For Each file In LoadFilesItemList
                    BackgroundWorker1.ReportProgress((filecount / LoadFilesItemList.Count) * 100)
                    Dim inputfile As String = file.FileName
                    Dim linecount As Integer = 0
                    Dim detectedEncoding As Encoding = Nothing

                    If file.FileEncodingType = "ASCII" Then
                        detectedEncoding = System.Text.Encoding.Default

                    ElseIf file.FileEncodingType = "BigEndianUnicode" Then
                        detectedEncoding = System.Text.Encoding.BigEndianUnicode

                    ElseIf file.FileEncodingType = "Unicode" Then
                        detectedEncoding = System.Text.Encoding.Unicode

                    ElseIf file.FileEncodingType = "UTF8" Then
                        detectedEncoding = System.Text.Encoding.UTF8
                    End If

                    Using fl As FileStream = New FileStream(inputfile, FileMode.Open, FileAccess.Read)
                        Using fl_in As StreamReader = New StreamReader(fl, detectedEncoding)

                            Dim line As String = fl_in.ReadLine
                            line = line.Replace(Chr(254), "")
                            If mainHeaderline = "" Then
                                mainHeaderline = line
                                Dim fields As String() = mainHeaderline.Split(Chr(20))

                                For Each fieldname In fields
                                    Dim fielditem As New MetaFieldItem
                                    fielditem.FieldName = fieldname
                                    fielditem.Value = ""
                                    fielditem.maxValue = ""
                                    fielditem.RecordCount = 0
                                    fielditem.Lenght = 0
                                    MetaFieldItemsList.Add(fielditem)
                                Next

                                MetadataFieldCount = MetaFieldItemsList.Count
                            End If

                            If mainHeaderline.ToLower = line.ToLower Then
                                While Not (fl_in.Peek() = -1)
                                    linecount += 1
                                    _globaltotallinecount += 1

                                    If _globaltotallinecount Mod 1000 = 0 Then
                                        BackgroundWorker1.ReportProgress((filecount / LoadFilesItemList.Count) * 100)
                                    End If

                                    line = fl_in.ReadLine
                                    line = line.Replace(Chr(254), "")
                                    Dim fielddata As String() = line.Split(Chr(20))

                                    If fielddata.Count <> MetadataFieldCount Then
                                        _globalerror = 1
                                        fstr_out.WriteLine(inputfile + "," + linecount.ToString + "," + "Record field count does Not match header column count.")
                                    Else

                                        For i = 0 To MetaFieldItemsList.Count - 1
                                            If MetaFieldItemsList.Item(i).Value = "" Then
                                                MetaFieldItemsList.Item(i).Value = fielddata(i).Trim
                                            End If

                                            If MetaFieldItemsList.Item(i).Lenght < fielddata(i).Count Then
                                                MetaFieldItemsList.Item(i).Lenght = fielddata(i).Count
                                                MetaFieldItemsList.Item(i).maxValue = fielddata(i).Trim
                                            End If

                                            If fielddata(i).Trim <> "" Then
                                                MetaFieldItemsList.Item(i).RecordCount += 1
                                            End If
                                        Next
                                    End If
                                End While
                            Else
                                _globalerror = 1
                                fstr_out.WriteLine(inputfile + "," + linecount.ToString + "," + "Header line does Not match group header line. Cannot analyze this file against the grouped files. Skipping file.")
                            End If
                        End Using
                    End Using
                    filecount += 1
                Next
            End Using
        End Using

        Using foutfile As FileStream = New FileStream(outputfile, FileMode.Create, FileAccess.Write)
            Using fstr_outfile As StreamWriter = New StreamWriter(foutfile, System.Text.Encoding.Unicode)

                fstr_outfile.WriteLine(Chr(34) + "**Notes: Quotes had been replaced with ^'s and tabs had been replaced with white space" + Chr(34))
                fstr_outfile.WriteLine(Chr(34) + "FieldName" + Chr(34) + "," + Chr(34) + "FirstEncounterValue(First 1000 chars)" + Chr(34) + "," + Chr(34) + "MaxEncounterValue(First 1000 chars)" + Chr(34) + "," + Chr(34) + "MaxFieldLenght" + Chr(34) + "," + Chr(34) + "NonBlankRecordCount" + Chr(34))

                For Each item In MetaFieldItemsList
                    Dim fieldvalue As String = item.Value
                    fieldvalue = fieldvalue.Replace(Chr(9), " ")
                    fieldvalue = fieldvalue.Replace(Chr(34), "^")

                    If fieldvalue.Count > 1000 Then
                        fieldvalue = Mid(fieldvalue, 1, 1000)
                    End If

                    Dim fieldmaxvalue As String = item.maxValue
                    fieldmaxvalue = fieldmaxvalue.Replace(Chr(9), " ")
                    fieldmaxvalue = fieldmaxvalue.Replace(Chr(34), "^")

                    If fieldmaxvalue.Count > 1000 Then
                        fieldmaxvalue = Mid(fieldmaxvalue, 1, 1000)
                    End If

                    fstr_outfile.WriteLine(Chr(34) + item.FieldName + Chr(34) + "," + Chr(34) + fieldvalue + Chr(34) + "," + Chr(34) + fieldmaxvalue + Chr(34) + "," + Chr(34) + item.Lenght.ToString + Chr(34) + "," + Chr(34) + item.RecordCount.ToString + Chr(34))
                Next
            End Using
        End Using

        BackgroundWorker1.ReportProgress((filecount / LoadFilesItemList.Count) * 100)
        Return 1
    End Function

    Function ReadLoadFileIndividually() As Integer
        Dim outdir As String = Path.GetDirectoryName(_OutFolder)
        Dim filecount As Integer = 1
        Using fout As FileStream = New FileStream(outdir + "\_Errors.Log", FileMode.Create, FileAccess.Write)
            Using fstr_out As StreamWriter = New StreamWriter(fout, System.Text.Encoding.UTF8)

                For Each file In LoadFilesItemList
                    BackgroundWorker1.ReportProgress((filecount / LoadFilesItemList.Count) * 100)

                    Dim inputfile As String = file.FileName
                    Dim outputfile As String = outdir + "\" + filecount.ToString + "_" + Path.GetFileNameWithoutExtension(inputfile) + "_report.csv"
                    Dim linecount As Integer = 0
                    Dim MetaFieldItemsList As New List(Of MetaFieldItem)
                    Dim MetadataFieldCount As Integer = 0
                    Dim detectedEncoding As Encoding = Nothing

                    If file.FileEncodingType = "ASCII" Then
                        detectedEncoding = System.Text.Encoding.Default

                    ElseIf file.FileEncodingType = "BigEndianUnicode" Then
                        detectedEncoding = System.Text.Encoding.BigEndianUnicode

                    ElseIf file.FileEncodingType = "Unicode" Then
                        detectedEncoding = System.Text.Encoding.Unicode

                    ElseIf file.FileEncodingType = "UTF8" Then
                        detectedEncoding = System.Text.Encoding.UTF8
                    End If

                    Using fl As FileStream = New FileStream(inputfile, FileMode.Open, FileAccess.Read)
                        Using fl_in As StreamReader = New StreamReader(fl, detectedEncoding)

                            Dim line As String = fl_in.ReadLine
                            line = line.Replace(Chr(254), "")

                            Dim fields As String() = line.Split(Chr(20))

                            For Each fieldname In fields
                                Dim fielditem As New MetaFieldItem
                                fielditem.FieldName = fieldname
                                fielditem.Value = ""
                                fielditem.maxValue = ""
                                fielditem.RecordCount = 0
                                fielditem.Lenght = 0
                                MetaFieldItemsList.Add(fielditem)
                            Next

                            MetadataFieldCount = MetaFieldItemsList.Count
                            While Not (fl_in.Peek() = -1)
                                linecount += 1
                                _globaltotallinecount += 1

                                If _globaltotallinecount Mod 1000 = 0 Then
                                    BackgroundWorker1.ReportProgress((filecount / LoadFilesItemList.Count) * 100)
                                End If

                                line = fl_in.ReadLine
                                line = line.Replace(Chr(254), "")
                                Dim fielddata As String() = line.Split(Chr(20))

                                If fielddata.Count <> MetadataFieldCount Then
                                    _globalerror = 1
                                    fstr_out.WriteLine(inputfile + "," + linecount.ToString + "," + "Record field count does not match header column count.")
                                Else
                                    For i = 0 To MetaFieldItemsList.Count - 1

                                        If MetaFieldItemsList.Item(i).Value = "" Then
                                            MetaFieldItemsList.Item(i).Value = fielddata(i).Trim
                                        End If

                                        If MetaFieldItemsList.Item(i).Lenght < fielddata(i).Count Then
                                            MetaFieldItemsList.Item(i).Lenght = fielddata(i).Count
                                            MetaFieldItemsList.Item(i).maxValue = fielddata(i).Trim
                                        End If

                                        If fielddata(i).Trim <> "" Then
                                            MetaFieldItemsList.Item(i).RecordCount += 1
                                        End If
                                    Next
                                End If
                            End While
                        End Using
                    End Using

                    Using foutfile As FileStream = New FileStream(outputfile, FileMode.Create, FileAccess.Write)
                        Using fstr_outfile As StreamWriter = New StreamWriter(foutfile, System.Text.Encoding.Unicode)

                            fstr_outfile.WriteLine(Chr(34) + "**Notes: Quotes had been replaced with ^'s and tabs had been replaced with white space" + Chr(34))
                            fstr_outfile.WriteLine(Chr(34) + "FieldName" + Chr(34) + "," + Chr(34) + "FirstEncounterValue(First 1000 chars)" + Chr(34) + "," + Chr(34) + "MaxEncounterValue(First 1000 chars)" + Chr(34) + "," + Chr(34) + "MaxFieldLenght" + Chr(34) + "," + Chr(34) + "NonBlankRecordCount" + Chr(34))

                            For Each item In MetaFieldItemsList
                                Dim fieldvalue As String = item.Value
                                fieldvalue = fieldvalue.Replace(Chr(9), " ")
                                fieldvalue = fieldvalue.Replace(Chr(34), "^")

                                If fieldvalue.Count > 1000 Then

                                    fieldvalue = Mid(fieldvalue, 1, 1000)
                                End If

                                Dim fieldmaxvalue As String = item.maxValue
                                fieldmaxvalue = fieldmaxvalue.Replace(Chr(9), " ")
                                fieldmaxvalue = fieldmaxvalue.Replace(Chr(34), "^")

                                If fieldmaxvalue.Count > 1000 Then

                                    fieldmaxvalue = Mid(fieldmaxvalue, 1, 1000)
                                End If
                                fstr_outfile.WriteLine(Chr(34) + item.FieldName + Chr(34) + "," + Chr(34) + fieldvalue + Chr(34) + "," + Chr(34) + fieldmaxvalue + Chr(34) + "," + Chr(34) + item.Lenght.ToString + Chr(34) + "," + Chr(34) + item.RecordCount.ToString + Chr(34))
                            Next
                        End Using
                    End Using

                    filecount += 1
                Next
            End Using
        End Using

        BackgroundWorker1.ReportProgress((filecount / LoadFilesItemList.Count) * 100)
        Return 1
    End Function

    Public Function GetFileEncoding(filePath As String) As Encoding
        Using sr As StreamReader = New StreamReader(filePath, True)
            sr.Read()
            Return sr.CurrentEncoding
        End Using
    End Function

    Public Function DetectEncodingFromBom(data() As Byte) As Encoding
        Dim detectedEncoding As Encoding = Nothing
        For Each info As EncodingInfo In Encoding.GetEncodings()
            Dim currentEncoding As Encoding = info.GetEncoding()
            Dim preamble() As Byte = currentEncoding.GetPreamble()
            Dim match As Boolean = True
            If (preamble.Length > 0) And (preamble.Length <= data.Length) Then
                For i As Integer = 0 To preamble.Length - 1
                    If preamble(i) <> data(i) Then
                        match = False
                        Exit For
                    End If
                Next
            Else
                match = False
            End If
            If match Then
                detectedEncoding = currentEncoding
                Exit For
            End If
        Next
        Return detectedEncoding
    End Function

    Public Function GetexternalFileName(line As String) As String
        Dim startpos As Integer = InStr(line, Chr(34)) + 1
        Dim endpos As Integer = InStrRev(line, Chr(34))
        Dim FileName As String = Mid(line, startpos, endpos - startpos)
        Return FileName
    End Function

    Private Sub Window_Closed(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Application.Current.Shutdown()
    End Sub

    Private Sub worker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        progressbar1.Value = e.ProgressPercentage
        LB_Count.Content = "Analyzing line: " + _globaltotallinecount.ToString
        Dim floor As Integer
        floor = Math.Floor(progressbar1.Value)
    End Sub

    Private Async Sub worker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Error IsNot Nothing Then
            progressbar1.Value = 0
            Await ShowMessageAsync("", "The application background thread encountered error during processing.")
            windowgrip.IsEnabled = True
        Else

            If e.Result <> "" Then
                windowgrip.IsEnabled = True

                Await ShowMessageAsync("", e.Result)
                progressbar1.Value = 0
            Else
                progressbar1.Value = 100

                If _processerror = 1 Then
                    Cursor = Cursors.Arrow
                    Await ShowMessageAsync("", "Error occured. Please correct error and rerun the tool.")
                Else

                    If _globalerror = 1 Then
                        Cursor = Cursors.Arrow
                        Await ShowMessageAsync("", "Complete with errors. Please check error log. ")

                    Else
                        Cursor = Cursors.Arrow
                        Await ShowMessageAsync("", "Complete! ")
                        Me.Height = "800"
                        CalculateFields()
                        Me.CalculateTab.IsSelected = True
                        Me.CalculateTab.IsEnabled = True
                        Me.Title = "LW Calculated Fields"
                        optionsTable.Visibility = System.Windows.Visibility.Visible
                        ' Me.Height = "1050"
                        ' CalculatedText1.Visibility = System.Windows.Visibility.Hidden
                    End If
                End If
                windowgrip.IsEnabled = True
            End If
        End If
    End Sub

    Private Sub BT_Output_Click(sender As Object, e As RoutedEventArgs) Handles BT_Output.Click
        Dim dlg As New Microsoft.Win32.SaveFileDialog
        dlg.InitialDirectory = Path.GetDirectoryName(Directory.GetCurrentDirectory)
        dlg.FileName = "Output.dat"
        dlg.DefaultExt = ".dat" ' Default file extension
        dlg.Filter = "Comma Delimited (.dat)|*.dat" ' Filter files by extension

        ' Show open file dialog box
        Dim result? As Boolean = dlg.ShowDialog()

        ' Process open file dialog box results
        If result = True Then
            ' Open document
            TB_Output.Text = dlg.FileName
        End If
    End Sub

    'Private Sub BT_XML_Click(sender As Object, e As RoutedEventArgs) Handles BT_XML.Click
    '    Dim dlg As New Microsoft.Win32.OpenFileDialog
    '    dlg.DefaultExt = ".dat" ' Default file extension
    '    dlg.Filter = "Dat Files (.dat)|*.dat" ' Filter files by extension
    '    'TB_XML.Text = ""

    '    ' Show open file dialog box
    '    dlg.Multiselect = True
    '    Dim result? As Boolean = dlg.ShowDialog()

    '    ' Process open file dialog box results
    '    If result = True Then
    '        ' Open document
    '        'TB_XML.Text = dlg.FileNames(0) + ";" + dlg.FileNames(1)
    '        'mainXMLFile = TB_XML.Text
    '        InitializeComponent()

    '        For Each item In dlg.FileNames
    '            Dim cb As New ComboBox

    '            cb.BorderThickness = New Thickness(0, 0, 0, 0)
    '            cb.Width = 142

    '            Dim tb As New TextBox
    '            Dim sb As New TextBox
    '            sb.BorderThickness = New Thickness(0, 0, 0, 0)
    '            tb.BorderThickness = New Thickness(0, 0, 0, 0)
    '            sb.IsReadOnly = True
    '            tb.IsReadOnly = True
    '            tb.Background = Brushes.Transparent
    '            tb.Text = item
    '            sb.Text = ""

    '            Dim name As String = item
    '            cb.Items.Add("ASCII")
    '            cb.Items.Add("BigEndianUnicode")
    '            cb.Items.Add("Unicode")
    '            cb.Items.Add("UTF8")

    '            Dim f As New FileItem
    '            f.FileName = tb
    '            f.Spacing = sb
    '            f.FileEncodingType = cb

    '            Dim detectedEncoding As Encoding = Nothing
    '            Dim arraySizeMinusOne = 6
    '            Dim buffer() As Byte = New Byte(arraySizeMinusOne) {}

    '            Using fs As New FileStream(f.FileName.Text, FileMode.Open, FileAccess.Read, FileShare.None)
    '                fs.Read(buffer, 0, buffer.Length)
    '            End Using

    '            ' Dim buffer() As Byte = File.ReadAllBytes(mainXMLFile)
    '            detectedEncoding = DetectEncodingFromBom(buffer)

    '            If detectedEncoding Is Nothing Then
    '                f.FileEncodingType.SelectedIndex = -1
    '            Else
    '                If detectedEncoding Is System.Text.Encoding.ASCII Then
    '                    f.FileEncodingType.SelectedIndex = 0

    '                ElseIf detectedEncoding Is System.Text.Encoding.BigEndianUnicode Then
    '                    f.FileEncodingType.SelectedIndex = 1

    '                ElseIf detectedEncoding Is System.Text.Encoding.Unicode Then
    '                    f.FileEncodingType.SelectedIndex = 2


    '                ElseIf detectedEncoding Is System.Text.Encoding.UTF8 Then
    '                    f.FileEncodingType.SelectedIndex = 3
    '                End If
    '            End If

    '            FileItemsList.Add(f)
    '        Next

    '        Dim newRow As RowDefinition
    '        For i = rowIndex To FileItemsList.Count - 1
    '            newRow = New RowDefinition
    '            newRow.Height = New GridLength(0, GridUnitType.Auto)

    '            gvProducts.RowDefinitions.Add(newRow)

    '            Grid.SetRow(FileItemsList.Item(i).FileName, rowIndex)
    '            Grid.SetColumn(FileItemsList.Item(i).FileName, 0)

    '            Grid.SetRow(FileItemsList.Item(i).Spacing, rowIndex)
    '            Grid.SetColumn(FileItemsList.Item(i).Spacing, 1)

    '            Grid.SetRow(FileItemsList.Item(i).FileEncodingType, rowIndex)
    '            Grid.SetColumn(FileItemsList.Item(i).FileEncodingType, 2)

    '            gvProducts.Children.Add(FileItemsList.Item(i).FileName)
    '            gvProducts.Children.Add(FileItemsList.Item(i).Spacing)
    '            gvProducts.Children.Add(FileItemsList.Item(i).FileEncodingType)
    '            rowIndex += 1
    '        Next
    '    End If
    'End Sub

    Private Sub ImagePanel_Drop(ByVal sender As Object, ByVal e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            ' Note that you can have more than one file.
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            ' Assuming you have one file that you care about, pass it off to whatever
            ' handling code you have defined.

            InitializeComponent()
            System.Array.Sort(Of String)(files)

            For Each fileitem In files
                If Path.GetExtension(fileitem).ToLower <> ".dat" Then
                    MsgBox(fileitem + " is not a .dat file. This file cannot be added.")

                ElseIf fileitem.Count > 256 Then
                    MsgBox(fileitem + " is long path. This file cannot be added.")
                Else
                    ' Open document
                    'TB_XML.Text = dlg.FileNames(0) + ";" + dlg.FileNames(1)
                    'mainXMLFile = TB_XML.Text

                    Dim cb As New ComboBox
                    cb.BorderThickness = New Thickness(0, 0, 0, 0)
                    cb.Items.Add("ASCII")
                    cb.Items.Add("BigEndianUnicode")
                    cb.Items.Add("Unicode")
                    cb.Items.Add("UTF8")

                    Dim tb As New TextBox
                    Dim sb As New TextBox
                    sb.BorderThickness = New Thickness(0, 0, 0, 0)
                    tb.BorderThickness = New Thickness(0, 0, 0, 0)
                    sb.IsReadOnly = True
                    tb.IsReadOnly = False
                    tb.Text = fileitem
                    sb.Text = ""

                    Dim name As String = fileitem
                    Dim f As New FileItem
                    f.FileName = tb
                    f.Spacing = sb
                    f.FileEncodingType = cb

                    Dim detectedEncoding As Encoding = Nothing
                    Dim arraySizeMinusOne = 6
                    Dim buffer() As Byte = New Byte(arraySizeMinusOne) {}

                    Using fs As New FileStream(f.FileName.Text, FileMode.Open, FileAccess.Read, FileShare.None)
                        fs.Read(buffer, 0, buffer.Length)
                    End Using

                    ' Dim buffer() As Byte = File.ReadAllBytes(mainXMLFile)
                    detectedEncoding = DetectEncodingFromBom(buffer)

                    If detectedEncoding Is Nothing Then
                        f.FileEncodingType.SelectedIndex = -1
                    Else

                        If detectedEncoding Is System.Text.Encoding.ASCII Then
                            f.FileEncodingType.SelectedIndex = 0

                        ElseIf detectedEncoding Is System.Text.Encoding.BigEndianUnicode Then
                            f.FileEncodingType.SelectedIndex = 1

                        ElseIf detectedEncoding Is System.Text.Encoding.Unicode Then
                            f.FileEncodingType.SelectedIndex = 2


                        ElseIf detectedEncoding Is System.Text.Encoding.UTF8 Then
                            f.FileEncodingType.SelectedIndex = 3
                        End If
                    End If

                    FileItemsList.Add(f)
                End If
            Next

            Dim newRow As RowDefinition
            'MsgBox(rowIndex.ToString + " " + (FileItemsList.Count - 1).ToString)
            For i = rowIndex To FileItemsList.Count - 1
                ' MsgBox("adding")
                newRow = New RowDefinition
                newRow.Height = New GridLength(0, GridUnitType.Auto)
                gvProducts.RowDefinitions.Add(newRow)

                Grid.SetRow(FileItemsList.Item(i).FileName, rowIndex)
                Grid.SetColumn(FileItemsList.Item(i).FileName, 0)

                Grid.SetRow(FileItemsList.Item(i).Spacing, rowIndex)
                Grid.SetColumn(FileItemsList.Item(i).Spacing, 1)

                Grid.SetRow(FileItemsList.Item(i).FileEncodingType, rowIndex)
                Grid.SetColumn(FileItemsList.Item(i).FileEncodingType, 2)

                gvProducts.Children.Add(FileItemsList.Item(i).FileName)
                gvProducts.Children.Add(FileItemsList.Item(i).Spacing)
                gvProducts.Children.Add(FileItemsList.Item(i).FileEncodingType)

                rowIndex += 1
            Next
        End If
    End Sub

    Function GetXMLEncoding() As String

        Dim log As String = mainXMLFile
        Dim fileEncoding As String = ""

        Using fl As FileStream = New FileStream(log, FileMode.Open, FileAccess.Read)
            Using fl_in As StreamReader = New StreamReader(fl, System.Text.Encoding.UTF8)

                If Not (fl_in.Peek() = -1) Then
                    Dim line As String = fl_in.ReadLine

                    If line.Contains("encoding=") Then
                        Dim startpos As Integer = InStr(line, "encoding=") + 10
                        Dim endpos As Integer = InStrRev(line, Chr(34))
                        fileEncoding = Mid(line, startpos, endpos - startpos)
                    End If
                End If
            End Using
        End Using

        If fileEncoding.ToUpper.Contains("UTF-16") = False And fileEncoding.ToUpper.Contains("UTF-8") = False Then
            Using fl As FileStream = New FileStream(log, FileMode.Open, FileAccess.Read)
                Using fl_in As StreamReader = New StreamReader(fl, System.Text.Encoding.Unicode)

                    If Not (fl_in.Peek() = -1) Then
                        Dim line As String = fl_in.ReadLine

                        If line.Contains("encoding=") Then
                            Dim startpos As Integer = InStr(line, "encoding=") + 10
                            Dim endpos As Integer = InStrRev(line, Chr(34))
                            fileEncoding = Mid(line, startpos, endpos - startpos)
                        End If
                    End If
                End Using
            End Using

        ElseIf fileEncoding = "" Then
            Using fl As FileStream = New FileStream(log, FileMode.Open, FileAccess.Read)
                Using fl_in As StreamReader = New StreamReader(fl, System.Text.Encoding.ASCII)

                    If Not (fl_in.Peek() = -1) Then
                        Dim line As String = fl_in.ReadLine

                        If line.Contains("encoding=") Then
                            Dim startpos As Integer = InStr(line, "encoding=") + 10
                            Dim endpos As Integer = InStrRev(line, Chr(34))

                            fileEncoding = Mid(line, startpos, endpos - startpos)
                        End If
                    End If
                End Using
            End Using
        End If
        'MsgBox(fileEncoding)
        Return fileEncoding.ToUpper
    End Function

    Private Sub CheckBox1_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox1 As CheckBox = TryCast(sender, CheckBox)
        pdExpander.Visibility = If(checkbox1.IsChecked, Visibility.Visible, Visibility.Collapsed)

        'If (checkbox1.IsChecked) Then
        '    parentDateGrid.Height = 200

        'End If

        'If Not (checkbox1.IsChecked) Then
        '    addMoreGrid.Visibility = Visibility.Collapsed
        '    parentDateGrid.Height = 0
        'End If
    End Sub

    Private Sub CheckBox2_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox2 As CheckBox = TryCast(sender, CheckBox)
        grExpander.Visibility = If(checkbox2.IsChecked, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub CheckBox3_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox3 As CheckBox = TryCast(sender, CheckBox)
        ftExpander.Visibility = If(checkbox3.IsChecked, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub CheckBox4_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox4 As CheckBox = TryCast(sender, CheckBox)
        edExpander.Visibility = If(checkbox4.IsChecked, Visibility.Visible, Visibility.Collapsed)


    End Sub

    Private Sub CheckBox5_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox5 As CheckBox = TryCast(sender, CheckBox)
        dfExpander.Visibility = If(checkBox5.IsChecked, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub CheckBox6_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox6 As CheckBox = TryCast(sender, CheckBox)
        cfExpander.Visibility = If(checkBox6.IsChecked, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub CheckBox7_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox7 As CheckBox = TryCast(sender, CheckBox)
        mfExpander.Visibility = If(checkbox7.IsChecked, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub CheckBox8_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox8 As CheckBox = TryCast(sender, CheckBox)
        cpcExpander.Visibility = If(checkbox8.IsChecked, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub CheckBox9_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox9 As CheckBox = TryCast(sender, CheckBox)
        sdExpander.Visibility = If(checkbox9.IsChecked, Visibility.Visible, Visibility.Collapsed)

        If (checkbox9.IsChecked) Then

            checkBox1.IsEnabled = False
            checkBox1.IsChecked = False
            pdExpander.Visibility = Visibility.Collapsed
            checkBox2.IsEnabled = False
            checkBox2.IsChecked = False
            grExpander.Visibility = Visibility.Collapsed
            checkBox3.IsEnabled = False
            checkBox3.IsChecked = False
            ftExpander.Visibility = Visibility.Collapsed
            checkBox4.IsEnabled = False
            checkBox4.IsChecked = False
            edExpander.Visibility = Visibility.Collapsed
            checkBox5.IsEnabled = False
            checkBox5.IsChecked = False
            dfExpander.Visibility = Visibility.Collapsed
            checkBox6.IsEnabled = False
            checkBox6.IsChecked = False
            cfExpander.Visibility = Visibility.Collapsed
            checkBox7.IsEnabled = False
            checkBox7.IsChecked = False
            mfExpander.Visibility = Visibility.Collapsed
            checkBox8.IsEnabled = False
            checkBox8.IsChecked = False
            cpcExpander.Visibility = Visibility.Collapsed
            checkBox10.IsEnabled = False
            checkBox10.IsChecked = False
            rcExpander.Visibility = Visibility.Collapsed
            checkBox11.IsEnabled = False
            checkBox11.IsChecked = False
            piciExpander.Visibility = Visibility.Collapsed
        End If

        If Not (checkbox9.IsChecked) Then
            checkBox1.IsEnabled = True
            checkBox2.IsEnabled = True
            checkBox3.IsEnabled = True
            checkBox4.IsEnabled = True
            checkBox5.IsEnabled = True
            checkBox6.IsEnabled = True
            checkBox7.IsEnabled = True
            checkBox8.IsEnabled = True
            checkBox10.IsEnabled = True
            checkBox11.IsEnabled = True
        End If
    End Sub

    Private Sub CheckBox10_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox10 As CheckBox = TryCast(sender, CheckBox)
        rcExpander.Visibility = If(checkbox10.IsChecked, Visibility.Visible, Visibility.Collapsed)

        If (checkbox10.IsChecked) Then
            checkBox1.IsEnabled = False
            checkBox1.IsChecked = False
            pdExpander.Visibility = Visibility.Collapsed
            checkBox2.IsEnabled = False
            checkBox2.IsChecked = False
            grExpander.Visibility = Visibility.Collapsed
            checkBox3.IsEnabled = False
            checkBox3.IsChecked = False
            ftExpander.Visibility = Visibility.Collapsed
            checkBox4.IsEnabled = False
            checkBox4.IsChecked = False
            edExpander.Visibility = Visibility.Collapsed
            checkBox5.IsEnabled = False
            checkBox5.IsChecked = False
            dfExpander.Visibility = Visibility.Collapsed
            checkBox6.IsEnabled = False
            checkBox6.IsChecked = False
            cfExpander.Visibility = Visibility.Collapsed
            checkBox7.IsEnabled = False
            checkBox7.IsChecked = False
            mfExpander.Visibility = Visibility.Collapsed
            checkBox8.IsEnabled = False
            checkBox8.IsChecked = False
            cpcExpander.Visibility = Visibility.Collapsed
            checkBox9.IsEnabled = False
            checkBox9.IsChecked = False
            sdExpander.Visibility = Visibility.Collapsed
            checkBox11.IsEnabled = False
            checkBox11.IsChecked = False
            piciExpander.Visibility = Visibility.Collapsed
        End If

        If Not (checkbox10.IsChecked) Then
            checkBox1.IsEnabled = True
            checkBox2.IsEnabled = True
            checkBox3.IsEnabled = True
            checkBox4.IsEnabled = True
            checkBox5.IsEnabled = True
            checkBox6.IsEnabled = True
            checkBox7.IsEnabled = True
            checkBox8.IsEnabled = True
            checkBox9.IsEnabled = True
            checkBox11.IsEnabled = True
        End If
    End Sub

    Private Sub CheckBox11_Changed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim checkbox11 As CheckBox = TryCast(sender, CheckBox)
        piciExpander.Visibility = If(checkbox11.IsChecked, Visibility.Visible, Visibility.Collapsed)

        If (checkbox11.IsChecked) Then
            checkBox1.IsEnabled = False
            checkBox1.IsChecked = False
            pdExpander.Visibility = Visibility.Collapsed
            checkBox2.IsEnabled = False
            checkBox2.IsChecked = False
            grExpander.Visibility = Visibility.Collapsed
            checkBox3.IsEnabled = False
            checkBox3.IsChecked = False
            ftExpander.Visibility = Visibility.Collapsed
            checkBox4.IsEnabled = False
            checkBox4.IsChecked = False
            edExpander.Visibility = Visibility.Collapsed
            checkBox5.IsEnabled = False
            checkBox5.IsChecked = False
            dfExpander.Visibility = Visibility.Collapsed
            checkBox6.IsEnabled = False
            checkBox6.IsChecked = False
            cfExpander.Visibility = Visibility.Collapsed
            checkBox7.IsEnabled = False
            checkBox7.IsChecked = False
            mfExpander.Visibility = Visibility.Collapsed
            checkBox8.IsEnabled = False
            checkBox8.IsChecked = False
            cpcExpander.Visibility = Visibility.Collapsed
            checkBox9.IsEnabled = False
            checkBox9.IsChecked = False
            sdExpander.Visibility = Visibility.Collapsed
            checkBox10.IsEnabled = False
            checkBox10.IsChecked = False
            rcExpander.Visibility = Visibility.Collapsed
        End If

        If Not (checkbox11.IsChecked) Then
            checkBox1.IsEnabled = True
            checkBox2.IsEnabled = True
            checkBox3.IsEnabled = True
            checkBox4.IsEnabled = True
            checkBox5.IsEnabled = True
            checkBox6.IsEnabled = True
            checkBox7.IsEnabled = True
            checkBox8.IsEnabled = True
            checkBox9.IsEnabled = True
            checkBox10.IsEnabled = True
        End If
    End Sub

    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Me.Closed
        closeProgramAlreadyRequested = False
        Me.Close()
    End Sub

    'Private Sub pdButton_Click3(sender As Object, e As EventArgs)
    '    hideParentDateSection.Visibility = System.Windows.Visibility.Visible
    '    showParentDateIcon.Visibility = System.Windows.Visibility.Visible
    '    hideParentDateIcon.Visibility = System.Windows.Visibility.Hidden
    'End Sub

    'Private Sub pdButton_Click3b(sender As Object, e As EventArgs)
    '    hideParentDateSection.Visibility = System.Windows.Visibility.Hidden
    '    showParentDateIcon.Visibility = System.Windows.Visibility.Hidden
    '    hideParentDateIcon.Visibility = System.Windows.Visibility.Visible
    'End Sub

    'Private Sub edButton_Click3(sender As Object, e As EventArgs)
    '    hideEmailDomainSection.Visibility = System.Windows.Visibility.Visible
    '    hideEmailDomainIcon.Visibility = System.Windows.Visibility.Hidden
    '    showEmailDomainIcon.Visibility = System.Windows.Visibility.Visible
    'End Sub

    'Private Sub edButton_Click3b(sender As Object, e As EventArgs)
    '    hideEmailDomainSection.Visibility = System.Windows.Visibility.Hidden
    '    showEmailDomainIcon.Visibility = System.Windows.Visibility.Hidden
    '    hideEmailDomainIcon.Visibility = System.Windows.Visibility.Visible
    'End Sub

    'Private Sub dtButton_Click3(sender As Object, e As EventArgs)
    '    hideDateTimeSection.Visibility = System.Windows.Visibility.Visible
    '    hideDateTimeIcon.Visibility = System.Windows.Visibility.Hidden
    '    showDateTimeIcon.Visibility = System.Windows.Visibility.Visible
    'End Sub

    'Private Sub dtButton_Click3b(sender As Object, e As EventArgs)
    '    hideDateTimeSection.Visibility = System.Windows.Visibility.Hidden
    '    showDateTimeIcon.Visibility = System.Windows.Visibility.Hidden
    '    hideDateTimeIcon.Visibility = System.Windows.Visibility.Visible
    'End Sub

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

    'Private Sub pdaddHierarchy_Click(sender As Object, e As RoutedEventArgs) Handles pdaddHierarchy.Click
    '    parentDateGrid.Height = 375
    '    addMoreGrid.Visibility = Visibility.Visible
    '    pdButtonClickCount += 1

    '    If (pdButtonClickCount = 1) Then
    '        pdAddMore1.Visibility = System.Windows.Visibility.Visible
    '        pdAddMore2.Visibility = System.Windows.Visibility.Visible
    '    End If

    '    If (pdButtonClickCount = 2) Then
    '        pdAddMore1b.Visibility = System.Windows.Visibility.Visible
    '        pdAddMore2b.Visibility = System.Windows.Visibility.Visible
    '    End If

    '    If (pdButtonClickCount = 3) Then
    '        pdAddMore1c.Visibility = System.Windows.Visibility.Visible
    '        pdAddMore2c.Visibility = System.Windows.Visibility.Visible
    '    End If

    '    If (pdButtonClickCount >= 4) Then
    '        If (pdButtonClickCount = 4) Then
    '            pdAddMore1d.Visibility = System.Windows.Visibility.Visible
    '            pdAddMore2d.Visibility = System.Windows.Visibility.Visible
    '        End If
    '        If (pdButtonClickCount = 5) Then
    '            pdAddMore1e.Visibility = System.Windows.Visibility.Visible
    '            pdAddMore2e.Visibility = System.Windows.Visibility.Visible
    '        End If
    '    End If

    '    If (pdButtonClickCount > 5) Then
    '        MsgBox("Reached maximum fields for this application")
    '    End If
    'End Sub

    'This function removes or replaces illegal characters from the file.
    Public Function RemoveIlegalChar(ByVal file As String) As String
        file = file.Replace("\\\", "\\")
        file = file.Replace(",", " |")
        Return file
    End Function

    Public Function CalculateFields() As Integer
        Dim filecount As Integer = 1
        Dim file As LoadFileItem = LoadFilesItemList(0)
        Dim inputfile As String = file.FileName
        Dim linecount As Integer = 0
        Dim detectedEncoding As Encoding = Nothing

        If file.FileEncodingType = "ASCII" Then
            detectedEncoding = System.Text.Encoding.Default

        ElseIf file.FileEncodingType = "BigEndianUnicode" Then
            detectedEncoding = System.Text.Encoding.BigEndianUnicode

        ElseIf file.FileEncodingType = "Unicode" Then
            detectedEncoding = System.Text.Encoding.Unicode

        ElseIf file.FileEncodingType = "UTF8" Then
            detectedEncoding = System.Text.Encoding.UTF8
        End If

        Using fl As FileStream = New FileStream(inputfile, FileMode.Open, FileAccess.Read)
            Using fl_in As StreamReader = New StreamReader(fl, detectedEncoding)

                Dim line As String = fl_in.ReadLine
                linecount += 1
                _globaltotallinecount += 1
                line = line.Replace(Chr(254), "")

                Dim columns As String() = line.Split(Chr(20))
                Dim columnList As List(Of column) = New List(Of column)

                Dim columnPosition As Integer = 0

                For Each c In columns

                    Dim column As column = New column


                    ' MsgBox(c)

                    column.ID = columnPosition
                    column.Name = c

                    columnList.Add(column)

                    columnPosition += 1
                Next


                globalBegDoc.Items.Clear()
                globalBegAttach.Items.Clear()

                'dictionary = New Dictionary(Of String, Integer)




                Dim cbList As List(Of ComboBox) = New List(Of ComboBox)

                cbList.Add(globalBegDoc)
                cbList.Add(globalBegAttach)
                cbList.Add(pdSentDate)
                cbList.Add(pdSentTime)
                cbList.Add(pdReceivedDate)
                cbList.Add(pdReceivedTime)
                cbList.Add(pdLastModDate)
                cbList.Add(pdLastModTime)
                cbList.Add(pdCreateDate)
                cbList.Add(pdCreateTime)
                cbList.Add(pdAccessDate)
                cbList.Add(pdAccessTime)
                cbList.Add(pdAddMore1)
                cbList.Add(pdAddMore1b)
                cbList.Add(pdAddMore1c)
                cbList.Add(pdAddMore1d)
                cbList.Add(pdAddMore1e)
                cbList.Add(pdAddMore2)
                cbList.Add(pdAddMore2b)
                cbList.Add(pdAddMore2c)
                cbList.Add(pdAddMore2d)
                cbList.Add(pdAddMore2e)

                cbList.Add(edFrom)
                cbList.Add(edTo)
                cbList.Add(edCC)
                cbList.Add(edBCC)
                cbList.Add(edSubject)
                cbList.Add(edFileType)

                For Each cb In cbList

                    Dim bind As Binding = New Binding
                    bind.Source = columnList


                    'globalBegDoc.ItemsSource = columnList



                    cb.DisplayMemberPath = "Name"
                    cb.SelectedValuePath = "ID"
                    cb.SetBinding(ComboBox.ItemsSourceProperty, bind)

                Next

                Dim lbBox As List(Of ListBox) = New List(Of ListBox)
                lbBox.Add(Listbox1)
                lbBox.Add(Listbox2)

                For Each lb In lbBox
                    Dim bind As Binding = New Binding
                    bind.Source = columnList

                    lb.DisplayMemberPath = "Name"
                    lb.SelectedValuePath = "ID"
                    lb.SetBinding(ListBox.ItemsSourceProperty, bind)

                Next

            End Using
        End Using
        filecount += 1
        Return 1
    End Function

    Private Function GetEmailDomains(ByVal begdoc As String, ByVal emails As String) As ArrayList
        'MsgBox(begdoc + ":" + emails)
        Dim source As String = emails
        Dim eamilDomains As ArrayList = New ArrayList
        Dim tempEamilDomains As ArrayList = New ArrayList
        If source <> "" Then
            Dim matches As MatchCollection


            matches = Regex.Matches(source, "@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})")


            For Each m In matches
                Dim tempDomain As String = Mid(m.value, 2)
                If tempEamilDomains.Contains(tempDomain.ToLower) = False Then
                    tempEamilDomains.Add(tempDomain.ToLower)
                    eamilDomains.Add(tempDomain)
                    ' MsgBox("Added" + ":" + begdoc + ":" + Mid(m.value, 2))
                End If

            Next
        End If


        Return eamilDomains
    End Function

    Private Sub Process(sender As Object, e As RoutedEventArgs)

        'MsgBox(globalBegDoc.SelectedValue.ToString)


        For Each file In LoadFilesItemList
            Dim records As List(Of Record) = New List(Of Record)
            Dim oinputfile As String = file.FileName
            Dim olinecount As Integer = 0
            Dim odetectedEncoding As Encoding = Nothing

            If file.FileEncodingType = "ASCII" Then
                odetectedEncoding = System.Text.Encoding.Default

            ElseIf file.FileEncodingType = "BigEndianUnicode" Then
                odetectedEncoding = System.Text.Encoding.BigEndianUnicode

            ElseIf file.FileEncodingType = "Unicode" Then
                odetectedEncoding = System.Text.Encoding.Unicode

            ElseIf file.FileEncodingType = "UTF8" Then
                odetectedEncoding = System.Text.Encoding.UTF8
            End If




            Using ofl As FileStream = New FileStream(oinputfile, FileMode.Open, FileAccess.Read)
                Using ofl_in As StreamReader = New StreamReader(ofl, odetectedEncoding)

                    Dim line As String = ofl_in.ReadLine
                    Dim currentRecordPosition As Integer = 1

                    While Not (ofl_in.Peek() = -1)
                        _globaltotallinecount += 1
                        line = ofl_in.ReadLine
                        line = line.Replace(Chr(254), "")

                        Dim ocolumn As String() = line.Split(Chr(20))
                        Dim oRecord As Record = New Record

                        oRecord.ID = currentRecordPosition

                        oRecord.BegDoc = Trim(ocolumn(globalBegDoc.SelectedValue))

                        If Trim(ocolumn(globalBegAttach.SelectedValue)) = "" Then
                            oRecord.BegAttach = Trim(ocolumn(globalBegDoc.SelectedValue))
                        Else
                            oRecord.BegAttach = Trim(ocolumn(globalBegAttach.SelectedValue))
                        End If


                        oRecord.DateTimeItems = New List(Of DateTimeItem)
#Region "ParentDate"
                        '1
                        If Not IsNothing(pdSentDate.SelectedItem) Then

                            If Trim(ocolumn(pdSentDate.SelectedValue)) <> "" Then


                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdSentDate.SelectedValue)

                                If Not IsNothing(pdSentTime.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdSentTime.SelectedValue)
                                End If

                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If

                        '2

                        If Not IsNothing(pdReceivedDate.SelectedItem) Then
                            If Trim(ocolumn(pdReceivedDate.SelectedValue)) <> "" Then

                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdReceivedDate.SelectedValue)

                                If Not IsNothing(pdReceivedTime.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdReceivedTime.SelectedValue)
                                End If
                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If

                        '3

                        If Not IsNothing(pdLastModDate.SelectedItem) Then
                            If Trim(ocolumn(pdLastModDate.SelectedValue)) <> "" Then

                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdLastModDate.SelectedValue)

                                If Not IsNothing(pdLastModTime.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdLastModTime.SelectedValue)
                                End If
                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If

                        '4
                        If Not IsNothing(pdCreateDate.SelectedItem) Then
                            If Trim(ocolumn(pdCreateDate.SelectedValue)) <> "" Then

                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdCreateDate.SelectedValue)

                                If Not IsNothing(pdCreateTime.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdCreateTime.SelectedValue)
                                End If
                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If

                        '5

                        If Not IsNothing(pdAccessDate.SelectedItem) Then
                            If Trim(ocolumn(pdAccessDate.SelectedValue)) <> "" Then

                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdAccessDate.SelectedValue)

                                If Not IsNothing(pdAccessTime.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdAccessTime.SelectedValue)
                                End If
                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If


                        '6
                        If Not IsNothing(pdAddMore1.SelectedItem) Then
                            If Trim(ocolumn(pdAddMore1.SelectedValue)) <> "" Then

                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdAddMore1.SelectedValue)

                                If Not IsNothing(pdAddMore2.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdAddMore2.SelectedValue)
                                End If
                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If


                        '7
                        If Not IsNothing(pdAddMore1b.SelectedItem) Then
                            If Trim(ocolumn(pdAddMore1b.SelectedValue)) <> "" Then

                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdAddMore1b.SelectedValue)

                                If Not IsNothing(pdAddMore2b.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdAddMore2b.SelectedValue)
                                End If
                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If

                        '8
                        If Not IsNothing(pdAddMore1c.SelectedItem) Then
                            If Trim(ocolumn(pdAddMore1c.SelectedValue)) <> "" Then

                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdAddMore1c.SelectedValue)

                                If Not IsNothing(pdAddMore2c.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdAddMore2c.SelectedValue)
                                End If
                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If

                        '9
                        If Not IsNothing(pdAddMore1d.SelectedItem) Then
                            If Trim(ocolumn(pdAddMore1d.SelectedValue)) <> "" Then

                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdAddMore1d.SelectedValue)

                                If Not IsNothing(pdAddMore2d.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdAddMore2d.SelectedValue)
                                End If
                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If

                        '10
                        If Not IsNothing(pdAddMore1e.SelectedItem) Then
                            If Trim(ocolumn(pdAddMore1e.SelectedValue)) <> "" Then

                                Dim DateTimeItem As DateTimeItem = New DateTimeItem
                                DateTimeItem.FieldDate = ocolumn(pdAddMore1e.SelectedValue)

                                If Not IsNothing(pdAddMore2e.SelectedItem) Then
                                    DateTimeItem.FieldTime = ocolumn(pdAddMore2e.SelectedValue)
                                End If
                                oRecord.DateTimeItems.Add(DateTimeItem)
                            End If
                        End If


                        If oRecord.DateTimeItems.Count > 0 Then
                            oRecord.DocDate = oRecord.DateTimeItems(0).FieldDate
                            oRecord.DocTime = oRecord.DateTimeItems(0).FieldTime
                        End If





                        'MsgBox(oRecord.DocDate + ":" + oRecord.DocTime)


#End Region

#Region "Email Domains Region"
                        'Process email domains

                        Dim eamilAdresses As String = ""
                        Dim oEmailDomains As ArrayList

                        If Not IsNothing(edFrom.SelectedItem) Then
                            If Trim(ocolumn(edFrom.SelectedValue)) <> "" Then

                                eamilAdresses = eamilAdresses + Trim(ocolumn(edFrom.SelectedValue))
                            End If

                        End If

                        If Not IsNothing(edTo.SelectedItem) Then
                            If Trim(ocolumn(edTo.SelectedValue)) <> "" Then

                                eamilAdresses = eamilAdresses + Trim(ocolumn(edTo.SelectedValue))
                            End If
                        End If

                        If Not IsNothing(edCC.SelectedItem) Then
                            If Trim(ocolumn(edCC.SelectedValue)) <> "" Then

                                eamilAdresses = eamilAdresses + Trim(ocolumn(edCC.SelectedValue))
                            End If
                        End If

                        If Not IsNothing(edBCC.SelectedItem) Then
                            If Trim(ocolumn(edBCC.SelectedValue)) <> "" Then

                                eamilAdresses = eamilAdresses + Trim(ocolumn(edBCC.SelectedValue))
                            End If
                        End If

                        If Not IsNothing(edSubject.SelectedItem) Then
                            If Trim(ocolumn(edSubject.SelectedValue)) <> "" Then

                                eamilAdresses = eamilAdresses + Trim(ocolumn(edSubject.SelectedValue))
                            End If
                        End If

                        If rb_internaldata.IsChecked Then

                            If Trim(eamilAdresses) = "" Then
                                If _nativeFileTypeList.Contains(Trim(ocolumn(edFileType.SelectedValue))) = True Then
                                    oRecord.Designation = "Internal"
                                Else
                                    oRecord.Designation = ""
                                End If
                            Else
                                oEmailDomains = GetEmailDomains(oRecord.BegDoc, eamilAdresses)

                                If oEmailDomains.Count > 0 Then

                                    Dim domainListStrings(oEmailDomains.Count - 1) As String

                                    oEmailDomains.CopyTo(domainListStrings)

                                    oRecord.EmailDomains = String.Join(";", domainListStrings)

                                    oRecord.Designation = "Internal"

                                    For Each d As String In oEmailDomains
                                        If _knownInternalDomains.Contains(d.ToLower) = False And d.ToLower.Contains("lw.com") = False Then
                                            oRecord.Designation = "External"
                                        End If

                                    Next

                                Else
                                    oRecord.Designation = "Internal"
                                End If

                            End If

                        Else
                            oEmailDomains = GetEmailDomains(oRecord.BegDoc, eamilAdresses)

                            Dim domainListStrings(oEmailDomains.Count - 1) As String

                            oEmailDomains.CopyTo(domainListStrings)

                            oRecord.EmailDomains = String.Join(";", domainListStrings)
                        End If








#End Region

                        currentRecordPosition += 1
                        records.Add(oRecord)
                    End While
                End Using
            End Using



            Dim sortedRecords As List(Of Record) = records.OrderBy(Function(r) r.BegAttach).ThenBy(Function(r) r.BegDoc).ToList

            For i = 0 To sortedRecords.Count - 1
                If i = 0 Then
                    sortedRecords(i).ParentDate = sortedRecords(i).DocDate
                    sortedRecords(i).ParentTime = sortedRecords(i).DocTime
                    sortedRecords(i).ParentChild = "P"
                Else

                    If sortedRecords(i).BegAttach = sortedRecords(i - 1).BegAttach Then
                        sortedRecords(i).ParentDate = sortedRecords(i - 1).ParentDate
                        sortedRecords(i).ParentTime = sortedRecords(i - 1).ParentTime
                        sortedRecords(i).ParentChild = "A"
                    Else
                        sortedRecords(i).ParentDate = sortedRecords(i).DocDate
                        sortedRecords(i).ParentTime = sortedRecords(i).DocTime
                        sortedRecords(i).ParentChild = "P"
                    End If
                End If



            Next


            sortedRecords.Sort(Function(x, y) x.ID.CompareTo(y.ID))

            'This is a temp location for testing Need to come back to modify
            Using foutfile As FileStream = New FileStream("\\lvdiprodata\EXPORTS01\PS100000\Erin Development Projects\LW Calculated Fields\testoutput.csv", FileMode.Create, FileAccess.Write)
                Using fstr_outfile As StreamWriter = New StreamWriter(foutfile, System.Text.Encoding.Unicode)

                    fstr_outfile.WriteLine("BegDoc,BegAttach,ParentDate,ParentTime,DocDate,DocTime,ParentChild,EmailDomains,Designation")

                    For Each r In sortedRecords

                        fstr_outfile.WriteLine(r.BegDoc + "," + r.BegAttach + "," + r.ParentDate + "," + r.ParentTime + "," + r.DocDate + "," + r.DocTime + "," + r.ParentChild + "," + r.EmailDomains + "," + r.Designation)


                    Next


                End Using
            End Using


        Next

        MsgBox("complete")


        'For Each r In records
        '    For Each dt In r.DateTimeItems
        '        MsgBox(r.BegDoc + "|" + r.BegAttach + "|" + dt.FieldDate + "|" + dt.FieldTime)
        '    Next
        'Next

        'For Each r In records
        '    For Each ed In r.EmailDomainItems
        '        MsgBox(r.BegDoc + "|" + r.BegAttach + "|" + ed.FieldEmailFrom + "|" + ed.FieldEmailTo + "|" + ed.FieldEmailCC + "|" + ed.FieldEmailBCC + "|" + ed.FieldEmailSubject)
        '    Next
        'Next

    End Sub

    Private Sub rb_internaldata_Checked(sender As Object, e As RoutedEventArgs) Handles rb_internaldata.Checked
        If rb_internaldata.IsChecked Then
            edFileType.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub rb_thirdparty_Checked(sender As Object, e As RoutedEventArgs) Handles rb_thirdparty.Checked
        If rb_thirdparty.IsChecked Then
            edFileType.Visibility = Visibility.Collapsed
        End If
    End Sub
End Class

