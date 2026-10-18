Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst


Public Class SolidWorksConfiguration

    Dim swApp As SldWorks '.SldWorks
    Dim swModel As ModelDoc2
    Dim propMgr As CustomPropertyManager
    Dim swEqnMgr As EquationMgr



    Public Sub Save(dataHubPath As String,
                    worksheetName As String)

        ' Referencia a SolidWorks

        'swApp = Marshal.GetActiveObject("SldWorks.Application")

        Try
            If swApp Is Nothing Then
                swApp = CreateObject("sldworks.application")
            Else
                swApp = GetObject("sldworks.application")
            End If
        Catch ex As Exception

        End Try

        swApp.Visible = True
        swModel = swApp.ActiveDoc

        If swModel Is Nothing Then

            MessageBox.Show("No active SolidWorks document found.")
            Exit Sub

        End If

        propMgr = swModel.Extension.CustomPropertyManager("")

        ' DataHub
        propMgr.Add3(
            "DataHub",
            swCustomInfoType_e.swCustomInfoText,
            dataHubPath,
            swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)

        ' Worksheet
        propMgr.Add3(
            "Worksheet",
            swCustomInfoType_e.swCustomInfoText,
            worksheetName,
            swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)

        MessageBox.Show("Properties successfully updated.")

    End Sub


    Public Sub SetParameters(parameterList As List(Of ParameterInfo))

        'swApp = Marshal.GetActiveObject("SldWorks.Application")

        Try
            If swApp Is Nothing Then
                swApp = CreateObject("sldworks.application")
            Else
                swApp = GetObject("sldworks.application")
            End If
        Catch ex As Exception

        End Try


        swModel = swApp.ActiveDoc

        If swModel Is Nothing Then Exit Sub

        swEqnMgr = swModel.GetEquationMgr()

        For Each parameter As ParameterInfo In parameterList

            'TODO:
            'Buscar variable global

            'Si existe
            'Actualizar

            'Si no existe
            'Crear

        Next

    End Sub
    'Private Function ParseComment(cadComment As String) As (ParameterCategory, String)

    '    Dim category As ParameterCategory
    '    Dim comment As String = ""

    '    If cadComment.StartsWith("Profile Parameter") Then

    '        category = ParameterCategory.Profile
    '        comment = cadComment.Replace("Profile Parameter", "").Trim()

    '    ElseIf cadComment.StartsWith("Key Parameter") Then

    '        category = ParameterCategory.Key
    '        comment = cadComment.Replace("Key Parameter", "").Trim()

    '    Else

    '        Throw New Exception("Unknown parameter category.")

    '    End If

    '    If comment.StartsWith("-") Then
    '        comment = comment.Substring(1).Trim()
    '    End If

    '    Return (category, comment)

    'End Function

    'Public Function GetParameters() As List(Of ParameterInfo)

    '    Dim parameters As New List(Of ParameterInfo)
    '    swModel = swApp.ActiveDoc
    '    swEqnMgr = swModel.GetEquationMgr()

    '    For i As Integer = 0 To swEqnMgr.GetCount - 1

    '        'Only process Global Variables
    '        If Not swEqnMgr.GlobalVariable(i) Then
    '            Continue For
    '        End If

    '        Dim equation As String = swEqnMgr.Equation(i)

    '        '-----------------------------------------
    '        ' Extract parameter name
    '        '-----------------------------------------

    '        Dim equalSignPosition As Integer = equation.IndexOf("=")

    '        If equalSignPosition = -1 Then
    '            Continue For
    '        End If

    '        Dim name As String = equation.Substring(0, equalSignPosition).Trim()

    '        'Remove quotation marks from the variable name
    '        name = name.Trim(""""c)

    '        '-----------------------------------------
    '        ' Extract value
    '        '-----------------------------------------

    '        Dim value As Double = swEqnMgr.Value(i)

    '        '-----------------------------------------
    '        ' Extract comment
    '        '-----------------------------------------

    '        Dim cadComment As String = ""

    '        Dim commentMarkerPosition As Integer = equation.IndexOf("'")

    '        If commentMarkerPosition >= 0 Then

    '            cadComment = equation.Substring(
    '            commentMarkerPosition + 1
    '        ).Trim()

    '        End If

    '        '-----------------------------------------
    '        ' Parse category and comment
    '        '-----------------------------------------

    '        Dim parsedComment = ParseComment(cadComment)

    '        If parsedComment.Item1 = ParameterCategory.None Then
    '            Continue For
    '        End If

    '        '-----------------------------------------
    '        ' Create ParameterInfo
    '        '-----------------------------------------

    '        Dim parameter As New ParameterInfo

    '        parameter.Name = name
    '        parameter.Value = value
    '        parameter.Category = parsedComment.Item1
    '        parameter.Comment = parsedComment.Item2

    '        parameters.Add(parameter)

    '    Next

    '    Return parameters

    'End Function

    Public Function GetConfiguration() As DataHubConfiguration


        Try
            If swApp Is Nothing Then
                swApp = CreateObject("sldworks.application")
            Else
                swApp = GetObject("sldworks.application")
            End If
        Catch ex As Exception

        End Try

        swApp.Visible = True
        swModel = swApp.ActiveDoc

        Dim config As New DataHubConfiguration

        propMgr =
            swModel.Extension.CustomPropertyManager("")

        Dim filePath As String = ""
        Dim worksheetName As String = ""

        Dim value As String = ""
        Dim resolvedValue As String = ""

        '-----------------------------------------
        ' Get DataHub filepath
        '-----------------------------------------

        propMgr.Get2("DataHub", value, resolvedValue)

        If resolvedValue <> "" Then
            filePath = resolvedValue
        End If

        '-----------------------------------------
        ' Get Worksheet
        '-----------------------------------------

        value = ""
        resolvedValue = ""

        propMgr.Get2("Worksheet", value, resolvedValue)

        If resolvedValue <> "" Then
            worksheetName = resolvedValue
        End If

        '-----------------------------------------
        ' Store configuration
        '-----------------------------------------

        config.FilePath = filePath
        config.WorksheetName = worksheetName

        Return config

    End Function


    Private Function ParseComment(cadComment As String) _
    As (ParameterCategory, ParameterMode, String)

        Dim category As ParameterCategory = ParameterCategory.None
        Dim mode As ParameterMode = ParameterMode.Design
        Dim comment As String = ""

        '-----------------------------------------
        ' Extract category
        '-----------------------------------------

        If cadComment.StartsWith("Profile Parameter") Then

            category = ParameterCategory.Profile

        ElseIf cadComment.StartsWith("Key Parameter") Then

            category = ParameterCategory.Key

        Else

            Return (ParameterCategory.None, ParameterMode.Design, "") ' Deshabilitar en caso que quiera llevarme params de categoría None

        End If

        '-----------------------------------------
        ' Remove category from string
        '-----------------------------------------

        Dim remaining As String = cadComment

        If category = ParameterCategory.Profile Then
            remaining = remaining.Substring("Profile Parameter".Length).Trim()
        ElseIf category = ParameterCategory.Key Then
            remaining = remaining.Substring("Key Parameter".Length).Trim()
        Else

        End If

        '-----------------------------------------
        ' Extract Mode
        '-----------------------------------------

        Dim modeStart As Integer = remaining.IndexOf("{")
        Dim modeEnd As Integer = remaining.IndexOf("}")

        If modeStart >= 0 AndAlso modeEnd > modeStart Then

            Dim modeCode As String =
                remaining.Substring(
                    modeStart + 1,
                    modeEnd - modeStart - 1
                ).Trim().ToUpper()

            Select Case modeCode

                Case "C"
                    mode = ParameterMode.Configuration

                Case "S"
                    mode = ParameterMode.Study

                Case "CS"
                    mode = ParameterMode.StudyAndConfiguration

                Case Else
                    mode = ParameterMode.Design

            End Select

            'Remove the {Mode} section
            remaining = remaining.Remove(
                modeStart,
                modeEnd - modeStart + 1
            ).Trim()

        End If

        '-----------------------------------------
        ' Extract Comment
        '-----------------------------------------

        If remaining.StartsWith("-") Then

            comment = remaining.Substring(1).Trim()

        Else

            comment = remaining.Trim()

        End If

        Return (category, mode, comment)

    End Function

    Public Function GetParameters() As List(Of ParameterInfo)

        Dim parameters As New List(Of ParameterInfo)

        swModel = swApp.ActiveDoc
        swEqnMgr = swModel.GetEquationMgr()

        '-----------------------------------------
        ' Document information
        '-----------------------------------------

        Dim location As String = swModel.GetPathName()

        Dim displayName As String =
            System.IO.Path.GetFileName(swModel.GetPathName())

        '-----------------------------------------
        ' Get Global Variables
        '-----------------------------------------

        For i As Integer = 0 To swEqnMgr.GetCount - 1

            'Only process Global Variables
            If Not swEqnMgr.GlobalVariable(i) Then
                Continue For
            End If

            Dim equation As String = swEqnMgr.Equation(i)

            '-----------------------------------------
            ' Extract parameter name
            '-----------------------------------------

            Dim equalSignPosition As Integer =
                equation.IndexOf("=")

            If equalSignPosition = -1 Then
                Continue For
            End If

            Dim name As String =
                equation.Substring(0, equalSignPosition).Trim()

            'Remove quotation marks from variable name
            name = name.Trim(""""c)

            '-----------------------------------------
            ' Extract value
            '-----------------------------------------

            Dim value As Double =
                swEqnMgr.Value(i)

            '-----------------------------------------
            ' Extract units
            '-----------------------------------------

            Dim units As String =
                GetParameterUnit(equation, equalSignPosition)

            '-----------------------------------------
            ' Extract CAD comment
            '-----------------------------------------

            Dim cadComment As String = ""

            Dim commentMarkerPosition As Integer =
                equation.IndexOf("'")

            If commentMarkerPosition >= 0 Then

                cadComment = equation.Substring(
                    commentMarkerPosition + 1
                ).Trim()

            End If

            '-----------------------------------------
            ' Parse category, mode and comment
            '-----------------------------------------

            Dim parsedComment =
                ParseComment(cadComment)

            '-----------------------------------------
            ' Create ParameterInfo
            '-----------------------------------------

            Dim parameter As New ParameterInfo

            parameter.Name = name
            parameter.Value = value
            parameter.Units = units
            parameter.Category = parsedComment.Item1
            parameter.Mode = parsedComment.Item2
            parameter.Comment = parsedComment.Item3

            parameter.Location = location
            parameter.DisplayName = displayName

            parameters.Add(parameter)

        Next

        Return parameters

    End Function

    Private Function GetParameterUnit(
    equation As String,
    equalSignPosition As Integer
) As String

        '-----------------------------------------
        ' Extract right side of equation
        '-----------------------------------------

        Dim rightSide As String =
            equation.Substring(equalSignPosition + 1).Trim()

        'Remove comment
        Dim commentPosition As Integer =
            rightSide.IndexOf("'")

        If commentPosition >= 0 Then

            rightSide =
                rightSide.Substring(0, commentPosition).Trim()

        End If

        '-----------------------------------------
        ' Check whether the expression is a
        ' reference to a SolidWorks dimension
        '-----------------------------------------
        rightSide = rightSide.Trim(""""c)

        'TEST --- BORRAR DESPUÉS

        'MsgBox("Right Side value (Potential Dimension) = " & rightSide)


        Dim dimension As Dimension = Nothing

        Try

            dimension = swModel.Parameter(rightSide)

        Catch

            dimension = Nothing

        End Try

        '-----------------------------------------
        ' Dimension reference found
        '-----------------------------------------

        If dimension IsNot Nothing Then
            'MsgBox("dimension type = " & dimension.GetType)
            Select Case dimension.GetType()

                Case 0
                    'Linear dimension
                    Return GetDefaultLinearUnit()

                Case 1
                    'Angular dimension
                    Return GetDefaultAngularUnit()

            End Select

        End If

        '-----------------------------------------
        ' Not a dimension reference
        '-----------------------------------------

        Return GetDefaultLinearUnit()

    End Function
    Private Function GetDefaultLinearUnit() As String

        Dim userUnit As UserUnit =
            swModel.GetUserUnit(
                CInt(swUserUnitsType_e.swLengthUnit)
            )

        Return userUnit.GetUnitsString(True)

    End Function
    Private Function GetDefaultAngularUnit() As String

        Select Case swEqnMgr.AngularEquationUnits

            Case 1
                Return "deg"

            Case 2
                Return "rad"

            Case Else
                Return ""

        End Select

    End Function
End Class
