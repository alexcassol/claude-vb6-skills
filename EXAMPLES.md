# Examples

Concrete before/after examples illustrating the conventions.

---

## 1. Case preservation in existing code

Adding a length check to an existing setter.

❌ **Wrong — "normalizing the case while I'm here":**

```vb
' Before:
Public Property Let Name(ByVal StrValue As String)
    m_StrName = StrValue
End Property

' After (diff shows every line changed):
Public Property Let Name(ByVal strValue As String)
    If Len(strValue) = 0 Then Exit Property
    m_strName = strValue
End Property
```

The diff now shows 3 changed lines. The reviewer has to inspect each one to
find the actual change.

✅ **Right — preserve existing case:**

```vb
' After:
Public Property Let Name(ByVal StrValue As String)
    If Len(StrValue) = 0 Then Exit Property
    m_StrName = StrValue
End Property
```

Diff shows 1 added line. The change is obvious at a glance.

---

## 2. Option Explicit and typed declarations

❌ **Before:**

```vb
Public Function CalculateTotal(items)
    Dim total
    Dim i
    For i = 0 To UBound(items)
        total = total + items(i)
    Next
    CalculateTotal = total
End Function
```

✅ **After:**

```vb
Option Explicit

Public Function CalculateTotal(ByRef items() As Double) As Double
    Dim total As Double
    Dim i As Long
    For i = LBound(items) To UBound(items)
        total = total + items(i)
    Next i
    CalculateTotal = total
End Function
```

Changes:

- `Option Explicit` at the top
- Return type declared (`As Double`)
- Parameter typed (`items() As Double`) and direction explicit (`ByRef`)
- Locals typed
- Loop counter `Long` (not `Integer` — would overflow at 32,768)
- `LBound`/`UBound` instead of assuming zero-based

---

## 3. On Error Resume Next — toxic in domain code

❌ **Before:**

```vb
On Error Resume Next
rs.Open sql, conn
Set product = New clsProduct
product.Load rs("id").Value
rs.Close
```

If `rs.Open` fails, `rs("id")` raises a new error that is also swallowed.
`product` is left in an invalid state. The user sees nothing.

✅ **After (CSEH ErrRaise):**

```vb
'CSEH: ErrRaise
Public Sub LoadProduct(ByVal lngID As Long)

        '<EhHeader>
        On Error GoTo LoadProduct_Err
        EnterMethod "modProduct", "LoadProduct"
        '</EhHeader>

        Dim objRs As Object
        Dim objProduct As clsProduct

100     InitializeDBConnection

105     Set objRs = CreateObject("ADODB.Recordset")
110     objRs.CursorLocation = adUseClient
115     Set objRs.ActiveConnection = mobjCnn
120     objRs.Open "SELECT * FROM Product WHERE idProduct = " & lngID, mobjCnn, adOpenForwardOnly, adLockReadOnly

125     If objRs.EOF Then Err.Raise vbObjectError + 101, "SampleApp.modProduct.LoadProduct", "Product not found: " & lngID

130     Set objProduct = New clsProduct
135     objProduct.Load objRs("idProduct").Value

    '<EhFooter>
        On Error GoTo 0

LoadProduct_Exit:
        If Not objRs Is Nothing Then
            If objRs.State = adStateOpen Then objRs.Close
            Set objRs = Nothing
        End If
        ExitMethod "modProduct", "LoadProduct"
        Exit Sub

LoadProduct_Err:
        Dim sLoadProduct_Err As String
        sLoadProduct_Err = "Error: " & Err.Number & " - " & Err.Description & " (" & Erl & ")"

        If Not objRs Is Nothing Then
            If objRs.State = adStateOpen Then objRs.Close
            Set objRs = Nothing
        End If
        ExitMethod "modProduct", "LoadProduct"

        Err.Raise vbObjectError + 100, "SampleApp.modProduct.LoadProduct", sLoadProduct_Err
    '</EhFooter>
End Sub
```

---

## 4. Public field vs Property

❌ **Before:**

```vb
' clsCustomer.cls
Public Name As String
Public Email As String
```

✅ **After:**

```vb
'--------------------------------------------------------------------------------
'    Component  : clsCustomer
'    Project    : SampleApp
'
'    Description: Domain object for a customer record.
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

' Variable to hold 'Name' property value
Private m_strName As String

' Variable to hold 'Email' property value
Private m_strEmail As String

Public Property Get Name() As String
    Name = m_strName
End Property

Public Property Let Name(ByVal strValue As String)
    m_strName = strValue
End Property

Public Property Get Email() As String
    Email = m_strEmail
End Property

Public Property Let Email(ByVal strValue As String)
    m_strEmail = strValue
End Property
```

The property form preserves binary compatibility for ActiveX DLLs, supports
data binding, and provides a place to add validation later without changing
callers.

**Note:** validation does **not** go in the setter. It goes in a named method
(`Validate`, `Load`, etc.) called explicitly by code that needs it.

---

## 5. Integer vs Long — silent overflow

❌ **Before:**

```vb
Dim counter As Integer
For counter = 1 To 50000  ' overflow at 32,768
    ' ...
Next counter
```

Runtime error 6 (Overflow) on iteration 32,768.

✅ **After:**

```vb
Dim counter As Long
For counter = 1 To 50000
    ' ...
Next counter
```

VB6 `Integer` is 16 bits (range -32,768 to 32,767). Almost always you want
`Long` (32 bits).

---

## 6. ByRef accidental mutation

❌ **Before:**

```vb
Public Sub Normalize(text As String)
    text = UCase(Trim(text))
End Sub

' Caller:
Dim name As String
name = "  john  "
Normalize name
' name is now "JOHN" — surprise for callers who did not expect mutation
```

VB6 defaults parameters to `ByRef`, so the caller's variable was modified
without obvious indication.

✅ **After — intent explicit:**

```vb
' If mutation is the intent:
Public Sub Normalize(ByRef text As String)
    text = UCase(Trim(text))
End Sub

' If returning a new string is the intent:
Public Function Normalize(ByVal text As String) As String
    Normalize = UCase(Trim(text))
End Function
```

---

## 7. SQL concatenation vs Parameters

❌ **Before — vulnerable to injection, breaks on apostrophes, locale-dependent:**

```vb
Public Function SearchCustomerByName(ByVal strName As String) As Object
    Dim strSQL As String
    Dim objRs As Object
    
    strSQL = "SELECT * FROM Customer " & _
             "WHERE custName LIKE '%" & strName & "%' " & _
             "  AND custCreatedAt > '" & Format(Date - 30, "dd/mm/yyyy") & "'"
    
    Set objRs = mobjCnn.Execute(strSQL)
    Set SearchCustomerByName = objRs
End Function
```

Problems:

- `strName = "O'Reilly"` breaks the SQL
- `strName = "'; DROP TABLE Customer; --"` is a SQL injection
- Date format depends on Windows regional settings — fails on machines with
  different locales
- Each unique `strName` creates a separate query plan in SQL Server's cache

✅ **After — parameterized:**

```vb
'CSEH: ErrRaise
Public Function SearchCustomerByName(ByVal strName As String) As Object

        '<EhHeader>
        On Error GoTo SearchCustomerByName_Err
        EnterMethod "modDB", "SearchCustomerByName"
        '</EhHeader>

        Dim objCmd As Object
        Dim objRs  As Object

100     InitializeDBConnection

105     Set objCmd = CreateObject("ADODB.Command")
110     Set objCmd.ActiveConnection = mobjCnn
115     objCmd.CommandType = adCmdText

120     objCmd.CommandText = "SELECT * FROM Customer " & _
                              "WHERE custName LIKE ? AND custCreatedAt > ?"

125     objCmd.Parameters.Append objCmd.CreateParameter("name", adVarChar, adParamInput, 100, "%" & strName & "%")
130     objCmd.Parameters.Append objCmd.CreateParameter("since", adDate, adParamInput, , Date - 30)

135     Set objRs = objCmd.Execute

140     Set SearchCustomerByName = objRs

    '<EhFooter>
        On Error GoTo 0

SearchCustomerByName_Exit:
        Set objCmd = Nothing
        ExitMethod "modDB", "SearchCustomerByName"
        Exit Function

SearchCustomerByName_Err:
        Dim sSearchCustomerByName_Err As String
        sSearchCustomerByName_Err = "Error: " & Err.Number & " - " & Err.Description & " (" & Erl & ")"

        Set objCmd = Nothing
        Set objRs = Nothing
        ExitMethod "modDB", "SearchCustomerByName"

        Err.Raise vbObjectError + 100, "SampleApp.modDB.SearchCustomerByName", sSearchCustomerByName_Err
    '</EhFooter>
End Function
```

---

## 8. Recordset leak

❌ **Before:**

```vb
Public Function GetCustomerName(ByVal lngID As Long) As String
    Dim objRs As Object
    Set objRs = mobjCnn.Execute("SELECT custName FROM Customer WHERE idCustomer=" & lngID)
    GetCustomerName = objRs("custName").Value
End Function
```

`objRs` is never closed or set to Nothing. In frequent calls (a loop, a
report), connection handles accumulate on the SQL Server.

✅ **After:**

```vb
'CSEH: ErrRaise
Public Function GetCustomerName(ByVal lngID As Long) As String

        '<EhHeader>
        On Error GoTo GetCustomerName_Err
        EnterMethod "modDB", "GetCustomerName"
        '</EhHeader>

        Dim objCmd As Object
        Dim objRs  As Object

100     InitializeDBConnection

105     Set objCmd = CreateObject("ADODB.Command")
110     Set objCmd.ActiveConnection = mobjCnn
115     objCmd.CommandType = adCmdText
120     objCmd.CommandText = "SELECT custName FROM Customer WHERE idCustomer = ?"
125     objCmd.Parameters.Append objCmd.CreateParameter("id", adInteger, adParamInput, , lngID)

130     Set objRs = objCmd.Execute

135     If Not objRs.EOF Then
140         GetCustomerName = fg_isNull(objRs("custName").Value, "")
        End If

    '<EhFooter>
        On Error GoTo 0

GetCustomerName_Exit:
        If Not objRs Is Nothing Then
            If objRs.State = adStateOpen Then objRs.Close
            Set objRs = Nothing
        End If
        Set objCmd = Nothing
        ExitMethod "modDB", "GetCustomerName"
        Exit Function

GetCustomerName_Err:
        Dim sGetCustomerName_Err As String
        sGetCustomerName_Err = "Error: " & Err.Number & " - " & Err.Description & " (" & Erl & ")"

        If Not objRs Is Nothing Then
            If objRs.State = adStateOpen Then objRs.Close
            Set objRs = Nothing
        End If
        Set objCmd = Nothing
        ExitMethod "modDB", "GetCustomerName"

        Err.Raise vbObjectError + 100, "SampleApp.modDB.GetCustomerName", sGetCustomerName_Err
    '</EhFooter>
End Function
```

---

## 9. Late binding for ADO

❌ **Before — locks the project to a specific MDAC version:**

```vb
' Requires reference to "Microsoft ActiveX Data Objects 2.8 Library"
Private mcnn As ADODB.Connection
Private mrs As ADODB.Recordset

Set mcnn = New ADODB.Connection
Set mrs = New ADODB.Recordset
```

Breaks on client machines with MDAC 6.0 or different versions.

✅ **After — works against whatever ADO is installed:**

```vb
Private mobjCnn As Object
Private mrsCustomers As Object

Set mobjCnn = CreateObject("ADODB.Connection")
Set mrsCustomers = CreateObject("ADODB.Recordset")
```

No type library reference. ADO constants declared locally:

```vb
Private Const adOpenForwardOnly As Long = 0
Private Const adLockReadOnly    As Long = 1
Private Const adUseClient       As Long = 3
```

---

## 10. EnterMethod / ExitMethod — both paths

❌ **Before — ExitMethod missing on the error path:**

```vb
Public Sub DoWork()
    On Error GoTo DoWork_Err
    EnterMethod "modX", "DoWork"
    
    ' ... work ...
    
    ExitMethod "modX", "DoWork"
    Exit Sub

DoWork_Err:
    LogMsg "Error: " & Err.Description
    ' ExitMethod missing — function stays "open" on the stack
End Sub
```

After this procedure raises, every subsequent `TraceMethod()` call shows
`modX::DoWork` as still being on the stack, which is wrong.

✅ **After — ExitMethod on both paths:**

```vb
'CSEH: ErrRaise
Public Sub DoWork()

        '<EhHeader>
        On Error GoTo DoWork_Err
        EnterMethod "modX", "DoWork"
        '</EhHeader>

        ' ... work ...

    '<EhFooter>
        On Error GoTo 0

DoWork_Exit:
        ExitMethod "modX", "DoWork"
        Exit Sub

DoWork_Err:
        Dim sDoWork_Err As String
        sDoWork_Err = "Error: " & Err.Number & " - " & Err.Description & " (" & Erl & ")"
        ExitMethod "modX", "DoWork"
        Err.Raise vbObjectError + 100, "SampleApp.modX.DoWork", sDoWork_Err
    '</EhFooter>
End Sub
```

---

## 11. MsgBox wrapper — not raw MsgBox

❌ **Before:**

```vb
If MsgBox("Save changes?", vbYesNo + vbQuestion) = vbYes Then
    ' ...
End If
```

Raw `MsgBox` does not integrate with the project's logging, does not match
the project's visual theme, and does not produce a consistent button layout
across forms.

✅ **After:**

```vb
If MsgBoxQuestion("Save changes?", , True) = vbYes Then
    ' ...
End If
```

The wrapper logs the prompt and the user's response automatically, displays
the project's custom dialog form, and centralizes the look-and-feel.

---

## 12. Line numbering for Erl

❌ **Before — Erl returns 0 in the handler:**

```vb
Public Sub ProcessOrder(ByVal lngID As Long)
    On Error GoTo ProcessOrder_Err
    
    Dim objCustomer As clsCustomer
    Set objCustomer = LoadCustomer(lngID)
    ValidateOrder objCustomer
    SubmitOrder objCustomer
    Exit Sub

ProcessOrder_Err:
    LogMsg "Error: " & Err.Description & " at line " & Erl
    ' Erl always returns 0 — lines are not numbered
End Sub
```

✅ **After — numbered, Erl is meaningful:**

```vb
'CSEH: ErrRaise
Public Sub ProcessOrder(ByVal lngID As Long)

        '<EhHeader>
        On Error GoTo ProcessOrder_Err
        EnterMethod "modOrder", "ProcessOrder"
        '</EhHeader>

        Dim objCustomer As clsCustomer

100     Set objCustomer = LoadCustomer(lngID)
105     ValidateOrder objCustomer
110     SubmitOrder objCustomer

    '<EhFooter>
        On Error GoTo 0

ProcessOrder_Exit:
        Set objCustomer = Nothing
        ExitMethod "modOrder", "ProcessOrder"
        Exit Sub

ProcessOrder_Err:
        Dim sProcessOrder_Err As String
        sProcessOrder_Err = "Error: " & Err.Number & " - " & Err.Description & " (" & Erl & ")"
        Set objCustomer = Nothing
        ExitMethod "modOrder", "ProcessOrder"
        Err.Raise vbObjectError + 100, "SampleApp.modOrder.ProcessOrder", sProcessOrder_Err
    '</EhFooter>
End Sub
```

Now `Erl` returns `100`, `105`, or `110` depending on where the failure
occurred — invaluable in production logs.
