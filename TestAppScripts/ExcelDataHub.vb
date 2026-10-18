Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Class ExcelDataHub

    Private _filePath As String
    Private _worksheetName As String

    Private _excelApp As Application
    Public _workbook As Workbook
    Public _worksheet As Worksheet

    Private Const COL_NAME As Integer = 1
    Private Const COL_VALUE As Integer = 2
    Private Const COL_UNITS As Integer = 3
    Private Const COL_CATEGORY As Integer = 4
    Private Const COL_MODE As Integer = 5
    Private Const COL_DESC As Integer = 6
    Private Const COL_DISPNAME As Integer = 7
    Private Const COL_PATH As Integer = 8


    Public ReadOnly Property FilePath As String
        Get
            Return _filePath
        End Get
    End Property

    Public ReadOnly Property WorksheetName As String
        Get
            Return _worksheetName
        End Get
    End Property

    Public Sub New(filePath As String,
                   worksheetName As String)

        _filePath = filePath
        _worksheetName = worksheetName

    End Sub


    ''' <summary>
    ''' Writes the parameters to the worksheet.
    ''' All-or-nothing: nothing is saved to disk unless every parameter was written.
    ''' Throws DataHubException on any failure.
    ''' </summary>
    Public Sub WriteParameters(parameterList As List(Of ParameterInfo))

        Try

            OpenWorkbook()

            Dim nextRow As Integer = GetNextAvailableRow()

            For Each parameter As ParameterInfo In parameterList

                Dim targetRow As Integer = FindParameterRow(parameter.Name)

                If targetRow = 0 Then
                    targetRow = nextRow
                    nextRow += 1
                End If

                WriteParameter(parameter, targetRow)

            Next

            ' Explicit save, only reached if every write succeeded
            _workbook.Save()

        Catch ex As COMException

            Throw New DataHubException(
                "Excel reported an error while writing the parameters: " & ex.Message, ex)

        Finally

            ' Always runs: never leaves a hidden Excel process or a locked file behind
            CloseWorkbook()

        End Try

    End Sub


    Public Function ReadProfileParameters() As List(Of ParameterInfo)

        Return New List(Of ParameterInfo)

    End Function


    Public Function ReadKeyParameters() As List(Of ParameterInfo)

        Return New List(Of ParameterInfo)

    End Function


    Public Sub Save()

    End Sub


    '-----------------------------------------
    ' Open / close
    '-----------------------------------------

    Private Sub OpenWorkbook()

        If Not System.IO.File.Exists(_filePath) Then
            Throw New DataHubException("The Data Hub file does not exist: " & _filePath)
        End If

        ' Checked BEFORE starting Excel, so no hidden dialog can ever block the application
        EnsureFileIsNotLocked()

        Try

            _excelApp = New Application
            _excelApp.Visible = False
            _excelApp.DisplayAlerts = False

            _workbook = _excelApp.Workbooks.Open(_filePath, UpdateLinks:=0)

            If _workbook.ReadOnly Then
                Throw New DataHubException(
                    "The Data Hub file was opened as read-only. Check the file attributes and permissions.")
            End If

            Try
                _worksheet = CType(_workbook.Worksheets(_worksheetName), Worksheet)
            Catch ex As Exception When TypeOf ex Is COMException OrElse TypeOf ex Is InvalidCastException
                Throw New DataHubException(
                    "The worksheet '" & _worksheetName & "' was not found in the Data Hub.", ex)
            End Try

        Catch ex As COMException

            Throw New DataHubException("Could not open the Data Hub with Excel: " & ex.Message, ex)

        End Try

    End Sub


    ''' <summary>
    ''' Excel keeps an open workbook locked, so a failed exclusive open means the file
    ''' is in use by another program (usually Excel itself).
    ''' </summary>
    Private Sub EnsureFileIsNotLocked()

        Try

            Using stream As New System.IO.FileStream(
                _filePath,
                System.IO.FileMode.Open,
                System.IO.FileAccess.ReadWrite,
                System.IO.FileShare.None)
            End Using

        Catch ex As System.IO.IOException

            Throw New DataHubException(
                "The Data Hub file is open in another program (probably Excel)." & vbCrLf &
                "Close it and try again.", ex)

        Catch ex As UnauthorizedAccessException

            Throw New DataHubException(
                "The Data Hub file cannot be modified (read-only file or missing permissions).", ex)

        End Try

    End Sub


    ''' <summary>
    ''' Safe to call at any point, even if the open was only partially completed.
    ''' Never saves (the explicit Save in WriteParameters is the only thing that persists changes)
    ''' and never throws, so it cannot hide the original error.
    ''' </summary>
    Private Sub CloseWorkbook()

        Try
            If _workbook IsNot Nothing Then _workbook.Close(False)
        Catch ex As Exception
            ' Cleanup must not mask the original error
        End Try

        Try
            If _excelApp IsNot Nothing Then _excelApp.Quit()
        Catch ex As Exception
            ' Cleanup must not mask the original error
        End Try

        ReleaseComObject(_worksheet)
        ReleaseComObject(_workbook)
        ReleaseComObject(_excelApp)

        _worksheet = Nothing
        _workbook = Nothing
        _excelApp = Nothing

        ' Releases the temporary COM wrappers created by chained calls like Cells(row, col).Value,
        ' which are the usual reason an EXCEL.EXE process stays alive after Quit
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub


    Private Sub ReleaseComObject(comObject As Object)

        If comObject Is Nothing Then Return

        Try
            Marshal.FinalReleaseComObject(comObject)
        Catch ex As Exception
            ' Cleanup must not mask the original error
        End Try

    End Sub


    '-----------------------------------------
    ' Row / cell helpers
    '-----------------------------------------

    Private Function FindParameterRow(parameterName As String) As Integer

        Dim row As Integer = 2

        While _worksheet.Cells(row, COL_NAME).Value IsNot Nothing

            If _worksheet.Cells(row, COL_NAME).Value.ToString = parameterName Then

                Return row

            End If

            row += 1

        End While

        Return 0

    End Function


    Private Sub WriteParameter(parameter As ParameterInfo,
                               row As Integer)

        _worksheet.Cells(row, COL_NAME).Value = parameter.Name

        _worksheet.Cells(row, COL_VALUE).Value = parameter.Value

        _worksheet.Cells(row, COL_CATEGORY).Value = parameter.Category.ToString

        _worksheet.Cells(row, COL_MODE).Value = parameter.Mode.ToString

        _worksheet.Cells(row, COL_DESC).Value = parameter.Comment

        _worksheet.Cells(row, COL_UNITS).Value = parameter.Units

        _worksheet.Cells(row, COL_DISPNAME).Value = parameter.DisplayName

        _worksheet.Cells(row, COL_PATH).Value = parameter.Location

    End Sub


    Private Function GetNextAvailableRow() As Integer

        Dim row As Integer = 2

        While _worksheet.Cells(row, COL_NAME).Value IsNot Nothing AndAlso
              _worksheet.Cells(row, COL_NAME).Value.ToString <> ""

            row += 1

        End While

        Return row

    End Function


    Public Sub SetFilePath(filePath As String)

        _filePath = filePath

    End Sub


    Public Sub SetWorksheetName(worksheetName As String)

        _worksheetName = worksheetName

    End Sub

End Class
