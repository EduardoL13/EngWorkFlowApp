Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Class ExcelDataHub

    'Private ReadOnly _filePath As String
    'Private ReadOnly _worksheetName As String

    Private _filePath As String
    Private _worksheetName As String


    Private _excelApp As Application
    'Private _workbook As Workbook
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


    Public Sub WriteParameters(parameterList As List(Of ParameterInfo))

        OpenWorkbook()

        Dim nextRow As Integer = GetNextAvailableRow()

        Try

            For Each parameter As ParameterInfo In parameterList

                Dim targetRow As Integer = FindParameterRow(parameter.Name)

                If targetRow = 0 Then
                    targetRow = nextRow
                    nextRow += 1
                End If

                WriteParameter(parameter, targetRow)

            Next

            _workbook.Save()

        Finally

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


    Private Sub OpenWorkbook()
        'Public Sub OpenWorkbook()

        _excelApp = New Application

        _excelApp.Visible = False

        _workbook = _excelApp.Workbooks.Open(_filePath)

        _worksheet = CType(_workbook.Worksheets(_worksheetName), Worksheet)

    End Sub


    Private Sub CloseWorkbook()
        'Public Sub CloseWorkbook()

        If _workbook IsNot Nothing Then
            _workbook.Close(True)
        End If

        If _excelApp IsNot Nothing Then
            _excelApp.Quit()
        End If

        Marshal.ReleaseComObject(_worksheet)
        Marshal.ReleaseComObject(_workbook)
        Marshal.ReleaseComObject(_excelApp)

        _worksheet = Nothing
        _workbook = Nothing
        _excelApp = Nothing

    End Sub


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
