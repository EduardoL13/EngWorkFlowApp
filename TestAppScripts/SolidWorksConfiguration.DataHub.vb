Imports System.Runtime.InteropServices
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

''' <summary>
''' Data Hub binding for SolidWorks: reads and writes the custom properties that
''' tell the system which Excel file and worksheet are the source of truth
''' for a given CAD document.
''' </summary>
Partial Public Class SolidWorksConfiguration

    ' Names of the custom properties stored in the CAD file.
    ' They are the contract between the CAD document and the rest of the system:
    ' do not rename them without a migration plan.
    Public Const DataHubPropertyName As String = "DataHub"
    Public Const WorksheetPropertyName As String = "Worksheet"

    ' Document captured by GetConfiguration (Initialize).
    ' Save always writes to this document, never to "whatever is active now".
    Private _targetDoc As ModelDoc2

    ''' <summary>
    ''' Title of the document captured at Initialize (empty if none).
    ''' </summary>
    Public ReadOnly Property TargetDocumentTitle As String
        Get
            If _targetDoc Is Nothing Then Return ""

            Try
                Return _targetDoc.GetTitle()
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Connects to a running SolidWorks session, or starts one if none is running.
    ''' Single place for connection logic (replaces the repeated CreateObject/GetObject blocks).
    ''' </summary>
    Private Sub ConnectToSolidWorks()

        ' Reuse the current connection if the session is still alive
        If swApp IsNot Nothing Then
            Try
                swApp.Visible = True
                Return
            Catch ex As Exception
                swApp = Nothing
            End Try
        End If

        Try
            ' Attach to a running SolidWorks session
            swApp = DirectCast(Marshal.GetActiveObject("SldWorks.Application"), SldWorks)
        Catch ex As COMException
            ' Not running: start a new session
            Try
                swApp = DirectCast(CreateObject("SldWorks.Application"), SldWorks)
            Catch startEx As Exception
                Throw New CadAdapterException("Could not connect to or start SolidWorks.", startEx)
            End Try
        End Try

        swApp.Visible = True

    End Sub

    ''' <summary>
    ''' Captures the active document as the target and reads its Data Hub properties.
    ''' Throws CadAdapterException if there is no active document.
    ''' </summary>
    Public Function GetConfiguration() As DataHubConfiguration

        ConnectToSolidWorks()

        Dim activeDoc As ModelDoc2 = TryCast(swApp.ActiveDoc, ModelDoc2)

        If activeDoc Is Nothing Then
            Throw New CadAdapterException("There is no active SolidWorks document.")
        End If

        _targetDoc = activeDoc
        swModel = activeDoc ' keep the shared field in sync for the other members

        Dim mgr As CustomPropertyManager = _targetDoc.Extension.CustomPropertyManager("")

        Dim config As New DataHubConfiguration
        config.FilePath = ReadProperty(mgr, DataHubPropertyName)
        config.WorksheetName = ReadProperty(mgr, WorksheetPropertyName)

        Return config

    End Function

    ''' <summary>
    ''' Writes the Data Hub properties to the document captured at Initialize.
    ''' Creates the properties if missing, replaces them otherwise.
    ''' The values are only persisted on disk once the document is saved (see SaveTargetDocument).
    ''' </summary>
    Public Sub Save(dataHubPath As String, worksheetName As String)

        EnsureTargetDocument()

        Dim mgr As CustomPropertyManager = _targetDoc.Extension.CustomPropertyManager("")

        WriteProperty(mgr, DataHubPropertyName, dataHubPath)
        WriteProperty(mgr, WorksheetPropertyName, worksheetName)

    End Sub

    ''' <summary>
    ''' Saves the target document so the custom properties persist on disk.
    ''' Returns False if the document has never been saved (no path) or the save failed.
    ''' </summary>
    Public Function SaveTargetDocument() As Boolean

        EnsureTargetDocument()

        If String.IsNullOrEmpty(_targetDoc.GetPathName()) Then Return False

        Dim errors As Integer = 0
        Dim warnings As Integer = 0

        Return _targetDoc.Save3(CInt(swSaveAsOptions_e.swSaveAsOptions_Silent), errors, warnings)

    End Function

    '-----------------------------------------
    ' Helpers
    '-----------------------------------------

    Private Function ReadProperty(mgr As CustomPropertyManager, propertyName As String) As String

        Dim rawValue As String = ""
        Dim resolvedValue As String = ""

        mgr.Get2(propertyName, rawValue, resolvedValue)

        Return If(resolvedValue, "")

    End Function

    Private Sub WriteProperty(mgr As CustomPropertyManager, propertyName As String, value As String)

        Dim result As Integer = mgr.Add3(
            propertyName,
            swCustomInfoType_e.swCustomInfoText,
            value,
            swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)

        If result <> CInt(swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged) Then
            Throw New CadAdapterException("Could not write the custom property '" & propertyName & "'.")
        End If

    End Sub

    Private Sub EnsureTargetDocument()

        If _targetDoc Is Nothing Then
            Throw New CadAdapterException("No document has been selected. Run Initialize first.")
        End If

        Try
            ' Any COM call fails if the document was closed since Initialize
            _targetDoc.GetTitle()
        Catch ex As Exception
            _targetDoc = Nothing
            Throw New CadAdapterException(
                "The document selected at Initialize is no longer available. Run Initialize again.", ex)
        End Try

    End Sub

End Class
