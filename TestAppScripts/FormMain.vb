Imports Microsoft.Office.Interop.Excel
Imports SolidWorks
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst 'Acceso al programa



Public Class FormMain

    Dim swApp As SldWorks
    Dim fullName As String
    Dim fileNameWithText As String
    Dim fileLocation As String
    Dim fileName As String
    Dim extension As String
    Private dataHub As ExcelDataHub
    Private adapter As SolidWorksConfiguration
    Private designFlow As DesignWorkflow






    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub startButton_Click(sender As Object, e As EventArgs) Handles startButton.Click
        'Try
        '    If swApp Is Nothing Then
        '        swApp = CreateObject("sldworks.application")
        '    Else
        '        swApp = GetObject("sldworks.application")
        '    End If
        'Catch ex As Exception

        'End Try

        'swApp.Visible = True

        ' ************************* NUEVA FUNCIÓN A INCORPORAR *********************************

        'Dim config As DataHubConfiguration = _solidWorksAdapter.GetConfiguration()

        'If config Is Nothing Then
        '    MessageBox.Show("No DataHub configuration found.")
        '    Return
        'End If

        'txtDataHub.Text = config.DataHubPath
        'txtWorksheet.Text = config.WorksheetName

    End Sub

    Private Sub infoDocButton_Click(sender As Object, e As EventArgs) Handles infoDocButton.Click
        Try
            If swApp Is Nothing Then
                swApp = CreateObject("sldworks.application")
            Else
                swApp = GetObject("sldworks.application")
            End If
        Catch ex As Exception

        End Try

        swApp.Visible = True


        Dim thisDoc As Object

        thisDoc = swApp.ActiveDoc   ' ACCEDER AQUI A UN PARAMETRO DEL EQNMGER PARA CONFIRMAR LO DEL PARSER

        Dim testDoc As ModelDoc2 = swApp.ActiveDoc
        'testDoc.fe

        Dim displayDim As SolidWorks.Interop.sldworks.Dimension 'D1@Sketch1

        'Dim displayDim2 As SolidWorks.Interop.sldworks.Dimension

        displayDim = testDoc.Parameter("D1@Sketch1")


        MsgBox("Valor parámetro es : " & displayDim.Value & "y es de tipo " & displayDim.GetType) '.SystemValue * 180 / 3.14)



        'Dim docEqnMger As EquationMgr = thisDoc.GetEquationMgr()
        'For i As Integer = 0 To docEqnMger.GetCount - 1
        '    MsgBox(docEqnMger.Equation(i) & " With Value: " & docEqnMger.Value(i))
        '    'docEqnMger.
        'Next

        'fullName = thisDoc.GetPathName 'Da nombre completo del archivo

        'If fullName <> "" Then
        '    fileLocation = System.IO.Path.GetDirectoryName(fullName)
        '    Debug.Print("Carpeta Documento: " & fileLocation)
        '    fileNameWithText = System.IO.Path.GetFileName(fullName)
        '    Debug.Print("fileNameWithText: " & fileNameWithText)
        '    fileName = System.IO.Path.GetFileNameWithoutExtension(fullName)
        '    Debug.Print("fileName: " & fileName)
        '    extension = System.IO.Path.GetExtension(fullName)
        '    Debug.Print("extension: " & extension)
        'End If

        'Dim tipoDoc As Integer

        'Select Case LCase(extension)
        '    Case ".sldprt"
        '        tipoDoc = 1
        '    Case ".sldasm"
        '        tipoDoc = 2
        '    Case ".slddrw"
        '        tipoDoc = 3
        'End Select

        'MsgBox("Documento es de tipo: " & tipoDoc)

    End Sub

    Private Sub BrowseButton_Click(sender As Object, e As EventArgs) Handles BrowseButton.Click
        Using dlg As New OpenFileDialog()

            dlg.Title = "Select Engineering DataHub"
            dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*"

            If dlg.ShowDialog() = DialogResult.OK Then

                txtDataHub.Text = dlg.FileName

            End If

        End Using
    End Sub

    Private Sub txtDataHub_TextChanged(sender As Object, e As EventArgs) Handles txtDataHub.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs) Handles loadDataDHButton.Click
        'Dim message As String

        'message =
        '    "DataHub:" & vbCrLf &
        '    txtDataHub.Text &
        '    vbCrLf & vbCrLf &
        '    "Worksheet:" & vbCrLf &
        '    txtWorksheet.Text

        'MessageBox.Show(message)


        adapter = New SolidWorksConfiguration()

        adapter.Save(txtDataHub.Text,
                     txtWorksheet.Text)
    End Sub

    Private Sub TestOpenDH_Click(sender As Object, e As EventArgs) Handles TestOpenDH.Click

        dataHub = New ExcelDataHub(txtDataHub.Text, txtWorksheet.Text)

        ' dataHub.OpenWorkbook()




    End Sub

    Private Sub TestCloseDH_Click(sender As Object, e As EventArgs) Handles TestCloseDH.Click

        ' dataHub.CloseWorkbook()

    End Sub

    Private Sub txtWorksheet_TextChanged(sender As Object, e As EventArgs) Handles txtWorksheet.TextChanged

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub SyncDW_Click(sender As Object, e As EventArgs) Handles SyncDW.Click

        'adapter = New SolidWorksConfiguration()
        dataHub = New ExcelDataHub(txtDataHub.Text, txtWorksheet.Text)

        designFlow = New DesignWorkflow(adapter, dataHub)
        designFlow.Run()

    End Sub

    Private Sub Initialize_Click(sender As Object, e As EventArgs) Handles Initialize.Click

        adapter = New SolidWorksConfiguration

        Dim config As DataHubConfiguration = adapter.GetConfiguration()

        Dim hasDataHub As Boolean = Not String.IsNullOrWhiteSpace(config.FilePath)
        Dim hasWorksheet As Boolean = Not String.IsNullOrWhiteSpace(config.WorksheetName)

        '-----------------------------------------
        ' Fill existing information
        '-----------------------------------------

        If hasDataHub Then
            txtDataHub.Text = config.FilePath
        Else
            txtDataHub.Clear()
        End If

        If hasWorksheet Then
            txtWorksheet.Text = config.WorksheetName
        Else
            txtWorksheet.Clear()
        End If

        '-----------------------------------------
        ' Notify user
        '-----------------------------------------

        If hasDataHub AndAlso hasWorksheet Then

            MessageBox.Show(
                "There is an existing Data Hub configuration in the current active document.",
                "Data Hub Configuration"
            )

        ElseIf hasDataHub AndAlso Not hasWorksheet Then

            MessageBox.Show(
                "Missing Worksheet property in the current active document.",
                "Data Hub Configuration"
            )

        ElseIf Not hasDataHub AndAlso hasWorksheet Then

            MessageBox.Show(
                "Missing Data Hub Filepath property in the current active document.",
                "Data Hub Configuration"
            )

        Else

            MessageBox.Show(
                "No existing Data Hub configuration properties in the current active document.",
                "Data Hub Configuration"
            )

        End If
    End Sub
End Class
