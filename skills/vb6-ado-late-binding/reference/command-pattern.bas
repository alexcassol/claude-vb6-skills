Attribute VB_Name = "modCommandPattern"
'--------------------------------------------------------------------------------
'    Component  : modCommandPattern
'    Project    : <YourProject>
'
'    Description: Helpers for the ADO Command + Parameters pattern with
'                 multibank support (MSSQL and PostgreSQL).
'
'                 - AddADOParameter / SetADOCommandParameters:
'                       accumulator pattern that buffers parameters in a
'                       Dictionary and applies them to a Command in one shot.
'
'                 - SetParameter / GetParameter:
'                       abstract the parameter naming convention difference
'                       between databases (MSSQL uses @name, Postgres uses
'                       in_name / out_name).
'
'                 Dependencies:
'                 - clsADOParameter (a small DTO class — see end of file)
'                 - modADOConstants (for adInteger, adParamInput, etc.)
'                 - A global gParameters object with a .DatabaseType property
'                   whose values match EnumDatabaseType.
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Public Enum EnumDatabaseType
    MSSQL = 1
    Postgres = 2
End Enum

' Buffer of parameters waiting to be applied to a Command
Private mdicPar As Dictionary

'---------------------------------------------------------------------------------------
' Procedure : AddADOParameter
' Purpose   : Adds or updates a parameter in the internal buffer. The buffer is
'             applied to a Command via SetADOCommandParameters.
'---------------------------------------------------------------------------------------
Public Sub AddADOParameter(ByVal ParName As String, _
                            ByVal ParType As Long, _
                            ByVal Value As Variant, _
                            Optional ByVal InputPar As Boolean = True)

    Dim objPar As clsADOParameter

    Set objPar = New clsADOParameter
    objPar.ParName = ParName
    objPar.ParType = ParType
    objPar.InputPar = InputPar
    objPar.Value = Value

    If mdicPar Is Nothing Then
        Set mdicPar = New Dictionary
    End If

    If mdicPar.Exists(ParName) Then
        Set mdicPar(ParName) = objPar
    Else
        mdicPar.Add ParName, objPar
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetADOCommandParameters
' Purpose   : Applies all buffered parameters to objCommand. By default clears
'             the buffer after applying.
'---------------------------------------------------------------------------------------
Public Sub SetADOCommandParameters(ByVal objCommand As Object, _
                                    Optional ByVal AfterClearParameters As Boolean = True)
    Const adParamInputOutput As Long = 3

    Dim i As Long
    Dim objPrm As Object
    Dim objP As clsADOParameter

    If mdicPar Is Nothing Then Exit Sub

    For i = 0 To mdicPar.Count - 1
        Set objP = mdicPar.Items(i)
        Set objPrm = objCommand.CreateParameter(objP.ParName, objP.ParType, adParamInputOutput, , objP.Value)
        objCommand.Parameters.Append objPrm
    Next i

    Set objP = Nothing
    Set objPrm = Nothing

    If AfterClearParameters Then
        Set mdicPar = Nothing
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetParameter
' Purpose   : Sets the value of a parameter already appended to objCommand,
'             using database-agnostic naming. Caller passes the bare name
'             (e.g., "idCustomer"); the helper prepends @ or in_/out_ based on
'             the active database type.
'---------------------------------------------------------------------------------------
Public Sub SetParameter(ByVal objCommand As Object, _
                         ByVal ParName As String, _
                         ByVal Value As Variant, _
                         Optional ByVal InputPar As Boolean = True)
    If gParameters.DatabaseType = Postgres Then
        ParName = IIf(InputPar, "in_", "out_") & ParName
    ElseIf gParameters.DatabaseType = MSSQL Then
        ParName = "@" & ParName
    End If
    objCommand.Parameters(ParName).Value = Value
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetParameter
' Purpose   : Reads the value of an output parameter from objCommand using
'             database-agnostic naming.
'---------------------------------------------------------------------------------------
Public Function GetParameter(ByVal objCommand As Object, _
                              ByVal ParName As String) As Variant
    If gParameters.DatabaseType = Postgres Then
        ParName = "out_" & ParName
    ElseIf gParameters.DatabaseType = MSSQL Then
        ParName = "@" & ParName
    End If
    GetParameter = objCommand.Parameters(ParName)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ShowADOCommand
' Purpose   : Renders a Command + its parameters as a SQL Server EXEC statement,
'             useful for copying and running directly in SSMS during debugging.
'---------------------------------------------------------------------------------------
Public Function ShowADOCommand(ByVal objCommand As Object) As String
    Dim strRet As String
    Dim lngCount As Long
    Dim i As Long
    Dim objPar As Object

    If objCommand Is Nothing Then
        ShowADOCommand = "Command is Nothing"
        Exit Function
    End If

    If objCommand.CommandText = "" Then
        ShowADOCommand = "CommandText is empty"
        Exit Function
    End If

    strRet = "EXEC dbo." & objCommand.CommandText & vbCrLf
    lngCount = objCommand.Parameters.Count

    If lngCount > 0 Then
        i = 0
        For Each objPar In objCommand.Parameters
            If i = 0 Then
                strRet = strRet & "    "
            Else
                strRet = strRet & "  , "
            End If
            strRet = strRet & objPar.Name & " = "

            If IsNull(objPar.Value) Then
                strRet = strRet & "NULL" & vbCrLf
            ElseIf IsNumeric(objPar.Value) Then
                strRet = strRet & objPar.Value & vbCrLf
            Else
                strRet = strRet & "'" & Replace(CStr(objPar.Value), "'", "''") & "'" & vbCrLf
            End If

            i = i + 1
        Next
    End If

    ShowADOCommand = strRet
End Function

'================================================================================
' Companion class: clsADOParameter.cls
'================================================================================
'
' Save the following as a separate .cls file with VB_Name = "clsADOParameter":
'
'   '--------------------------------------------------------------------------------
'   '    Component  : clsADOParameter
'   '    Project    : <YourProject>
'   '
'   '    Description: Small DTO holding parameter metadata until it is applied
'   '                 to an ADODB.Command.
'   '--------------------------------------------------------------------------------
'   Option Explicit
'
'   ' Variable to hold 'ParName' property value
'   Private m_strParName As String
'
'   ' Variable to hold 'ParType' property value
'   Private m_lngParType As Long
'
'   ' Variable to hold 'Value' property value
'   Private m_varValue As Variant
'
'   ' Variable to hold 'InputPar' property value
'   Private m_blnInputPar As Boolean
'
'   Public Property Get ParName() As String
'       ParName = m_strParName
'   End Property
'   Public Property Let ParName(ByVal strValue As String)
'       m_strParName = strValue
'   End Property
'
'   Public Property Get ParType() As Long
'       ParType = m_lngParType
'   End Property
'   Public Property Let ParType(ByVal lngValue As Long)
'       m_lngParType = lngValue
'   End Property
'
'   Public Property Get Value() As Variant
'       If IsObject(m_varValue) Then
'           Set Value = m_varValue
'       Else
'           Value = m_varValue
'       End If
'   End Property
'   Public Property Let Value(ByVal varValue As Variant)
'       m_varValue = varValue
'   End Property
'
'   Public Property Get InputPar() As Boolean
'       InputPar = m_blnInputPar
'   End Property
'   Public Property Let InputPar(ByVal blnValue As Boolean)
'       m_blnInputPar = blnValue
'   End Property
