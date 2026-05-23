Attribute VB_Name = "modADOConstants"
'--------------------------------------------------------------------------------
'    Component  : modADOConstants
'    Project    : <YourProject>
'
'    Description: ADO constants for late-binding scenarios where the
'                 "Microsoft ActiveX Data Objects" type library is not
'                 referenced. Declare Public so they are visible to every
'                 module that imports this file.
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

' -- Command Types ----------------------------------------------------------
Public Const adCmdUnknown        As Long = 8
Public Const adCmdText           As Long = 1
Public Const adCmdTable          As Long = 2
Public Const adCmdStoredProc     As Long = 4
Public Const adCmdFile           As Long = 256
Public Const adCmdTableDirect    As Long = 512

' -- Cursor Types -----------------------------------------------------------
Public Const adOpenUnspecified   As Long = -1
Public Const adOpenForwardOnly   As Long = 0
Public Const adOpenKeyset        As Long = 1
Public Const adOpenDynamic       As Long = 2
Public Const adOpenStatic        As Long = 3

' -- Cursor Locations -------------------------------------------------------
Public Const adUseNone           As Long = 1
Public Const adUseServer         As Long = 2
Public Const adUseClient         As Long = 3

' -- Lock Types -------------------------------------------------------------
Public Const adLockUnspecified   As Long = -1
Public Const adLockReadOnly      As Long = 1
Public Const adLockPessimistic   As Long = 2
Public Const adLockOptimistic    As Long = 3
Public Const adLockBatchOptimistic As Long = 4

' -- Connection / Object States --------------------------------------------
Public Const adStateClosed       As Long = 0
Public Const adStateOpen         As Long = 1
Public Const adStateConnecting   As Long = 2
Public Const adStateExecuting    As Long = 4
Public Const adStateFetching     As Long = 8

' -- Parameter Directions ---------------------------------------------------
Public Const adParamUnknown      As Long = 0
Public Const adParamInput        As Long = 1
Public Const adParamOutput       As Long = 2
Public Const adParamInputOutput  As Long = 3
Public Const adParamReturnValue  As Long = 4

' -- Data Types (common subset) --------------------------------------------
Public Const adEmpty             As Long = 0
Public Const adSmallInt          As Long = 2
Public Const adInteger           As Long = 3
Public Const adSingle            As Long = 4
Public Const adDouble            As Long = 5
Public Const adCurrency          As Long = 6
Public Const adDate              As Long = 7
Public Const adBSTR              As Long = 8
Public Const adIDispatch         As Long = 9
Public Const adError             As Long = 10
Public Const adBoolean           As Long = 11
Public Const adVariant           As Long = 12
Public Const adIUnknown          As Long = 13
Public Const adDecimal           As Long = 14
Public Const adTinyInt           As Long = 16
Public Const adUnsignedTinyInt   As Long = 17
Public Const adUnsignedSmallInt  As Long = 18
Public Const adUnsignedInt       As Long = 19
Public Const adBigInt            As Long = 20
Public Const adUnsignedBigInt    As Long = 21
Public Const adGUID              As Long = 72
Public Const adBinary            As Long = 128
Public Const adChar              As Long = 129
Public Const adWChar             As Long = 130
Public Const adNumeric           As Long = 131
Public Const adDBDate            As Long = 133
Public Const adDBTime            As Long = 134
Public Const adDBTimeStamp       As Long = 135
Public Const adVarChar           As Long = 200
Public Const adLongVarChar       As Long = 201
Public Const adVarWChar          As Long = 202
Public Const adLongVarWChar      As Long = 203
Public Const adVarBinary         As Long = 204
Public Const adLongVarBinary     As Long = 205

' -- Execute Options --------------------------------------------------------
Public Const adAsyncExecute      As Long = 16
Public Const adAsyncFetch        As Long = 32
Public Const adExecuteNoRecords  As Long = 128
Public Const adExecuteStream     As Long = 1024

' -- Transaction Isolation Levels ------------------------------------------
Public Const adXactReadUncommitted As Long = 256
Public Const adXactReadCommitted   As Long = 4096
Public Const adXactRepeatableRead  As Long = 65536
Public Const adXactIsolated        As Long = 1048576

' -- GetString Format ------------------------------------------------------
Public Const adClipString        As Long = 2
