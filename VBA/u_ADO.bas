Attribute VB_Name = "u_ADO"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' version 2015-06-17
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public g_dbConnection  As New ADODB.Connection
    Public g_strConnection As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Function _
Open_Client( _
            dbsConnection As ADODB.Connection, _
            strConString As String, _
            Optional strErr As String = "", _
            Optional timeout As Integer = 120) _
As Boolean
    On Error GoTo Err_
    Dim blnRetVal As Boolean
    
    Set dbsConnection = New ADODB.Connection
    With dbsConnection
        .CursorLocation = adUseClient
        .Provider = "MSDASQL"
        .CommandTimeout = timeout
        .Open (strConString)
    End With
    
    blnRetVal = True
    GoTo Exit_
    
Err_:
    blnRetVal = False
    strErr = Err.Description
    
Exit_:
    Open_Client = blnRetVal
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function _
Execute_SQL( _
            dbsConnection As ADODB.Connection, _
            rstRecordset As ADODB.Recordset, _
            strSQL As String, _
            Optional strErr As String = "") _
As Boolean
    On Error GoTo Err_
    Dim blnRetVal As Boolean
    
    With rstRecordset
        .CursorType = adOpenDynamic
        .Source = strSQL
        .LockType = adLockOptimistic
        .CacheSize = 100
        Set .ActiveConnection = dbsConnection
        .Open
    End With
    
    blnRetVal = True
    GoTo Exit_
    
Err_:
    strErr = Err.Description
    blnRetVal = False
    
Exit_:
    Execute_SQL = blnRetVal
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function _
rngExecSQL _
           (strConn As String, _
            SQLquery As String, _
            Optional blnShowHeadings As Boolean = True, _
            Optional blnReuseConnection As Boolean = True) _
As Variant
    Dim varRV
    On Error GoTo Err_
    
    Dim strErr As String
    If blnReuseConnection Then
        If Not g_dbConnection.State = adStateOpen Then
            If Not Open_Client(g_dbConnection, strConn, strErr) Then GoTo Err_
        End If
    Else
        If g_dbConnection.State = adStateOpen Then
            g_dbConnection.Close
            If Not Open_Client(g_dbConnection, strConn, strErr) Then GoTo Err_
        End If
    End If
    
    Dim rstData As New ADODB.Recordset
    If Execute_SQL(g_dbConnection, rstData, SQLquery, strErr) Then
        varRV = RecordSet2Array(rstData, blnShowHeadings)
    Else
        varRV = strErr
    End If
              
    GoTo Exit_
    
Err_:
    MsgBox "rngExecSQL error: " & Err.Description, vbExclamation, "Error"
    
Exit_:
    rngExecSQL = varRV
    
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function _
CallSqlFunc _
           (strConn As String, _
            SQLquery As String, _
            sqlPARAM As String, _
            SQLParamValue As String, _
            Optional blnShowHeadings As Boolean = False) _
As Variant
    Dim varRV
    On Error GoTo Err_
    
    Dim strErr As String
    If Not g_dbConnection.State = adStateOpen Then
        If Not Open_Client(g_dbConnection, strConn, strErr) Then GoTo Err_
    End If
    
    Dim strSQL As String
    Dim rstData As New ADODB.Recordset
    strSQL = Replace(SQLquery, sqlPARAM, SQLParamValue)
    If Execute_SQL(g_dbConnection, rstData, strSQL, strErr) Then
        varRV = RecordSet2Array(rstData, blnShowHeadings)
    End If
              
    GoTo Exit_
    
Err_:
    MsgBox "CallSqlFunc error: " & Err.Description, vbExclamation, "Error"
    
Exit_:
    CallSqlFunc = varRV
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function _
SqlMultiParamFunc _
           (strConn As String, _
            SQLquery As String, _
            rngParamNames As Range, _
            rngParamValues As Range, _
            Optional blnShowHeadings As Boolean = False) _
As Variant
    Dim varRV
    On Error GoTo Err_
    
    Dim strErr As String
    If Not g_dbConnection.State = adStateOpen Then
        If Not Open_Client(g_dbConnection, strConn, strErr) Then GoTo Err_
    End If
    
    Dim strPname As String
    Dim ParamVal
    Dim strPval  As String
    
    
    Dim rstData As New ADODB.Recordset
    Dim row As Integer
    row = 1
    Do While rngParamNames(row, 1) <> ""
        strPname = rngParamNames(row, 1)
        strPval = rngParamValues(row, 1)
        SQLquery = Replace(SQLquery, strPname, strPval)
        row = row + 1
    Loop
    
    If Execute_SQL(g_dbConnection, rstData, SQLquery, strErr) Then
        varRV = RecordSet2Array(rstData, blnShowHeadings)
    End If
              
    GoTo Exit_
    
Err_:
    MsgBox "CallSqlFunc error: " & Err.Description, vbExclamation, "Error"
    
Exit_:
    SqlMultiParamFunc = varRV
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function _
RecordSet2ExcelRangeViaArrays( _
            records As Recordset, _
            rngTgt As Range, _
            Optional showHeadings As Boolean = True, _
            Optional filterOn As Boolean = False) _
As Boolean
    Dim retVal As Boolean
    Dim showError As Boolean
    On Error GoTo Exit_
    
    Dim rngW As Range
    Set rngW = rngTgt
    
    Dim fldCount As Integer
    fldCount = records.Fields.Count
    If showHeadings Then
        Dim colNames() As String
        ReDim colNames(1, fldCount)
        
        Dim col As Integer
        For col = 0 To fldCount - 1
            colNames(0, col) = records.Fields(col).Name
        Next
        rngW.Resize(1, fldCount) = colNames
        Set rngW = rngTgt.Offset(1)
    End If
        
    showError = True
    Dim recArray As Variant
    recArray = records.GetRows
    Dim nrRows As Integer
    nrRows = UBound(recArray, 2) + 1
    rngW.Resize(nrRows, fldCount) = TransposeDim(recArray)

    
    If filterOn Then
        rngW.AutoFilter
    End If
    
    retVal = True
    
Exit_:
    If showError And retVal = False Then
        MsgBox "RecordSet2ExcelRangeViaArrays:" & vbCrLf & Err.Description, vbExclamation, "Error"
    End If
    RecordSet2ExcelRangeViaArrays = retVal
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function _
RecordSet2Array( _
            records As Recordset, _
            Optional showHeadings As Boolean = True) _
As Variant
    Dim retVal() As Variant
    On Error GoTo Exit_
    
    Dim recArray As Variant
    recArray = records.GetRows
    Dim nrRows As Integer
    nrRows = UBound(recArray, 2) + 1
    recArray = TransposeDim(recArray)

    If showHeadings Then
        Dim fldCount As Integer
        fldCount = records.Fields.Count
        ReDim retVal(0 To nrRows, 0 To fldCount - 1)
        Dim col As Integer
        Dim row As Integer
        For col = 0 To fldCount - 1
            retVal(0, col) = records.Fields(col).Name
        Next
        For row = 1 To nrRows
            For col = 0 To fldCount - 1
                retVal(row, col) = recArray(row - 1, col)
            Next col
        Next row
    Else
        retVal = recArray
    End If

    
    
Exit_:
    RecordSet2Array = retVal
    
End Function

' inlcude only those recrods with a string in column intCol listed in astrFilter
'   when blnIncludeFilter is true,
'   otherwise exclude them
Public Function _
FilteredRecordSet2Array( _
            records As Recordset, _
            astrFilter() As String, _
            intCol As Integer, _
            Optional showHeadings As Boolean = True, _
            Optional blnIncludeFilter As Boolean = True) _
As Variant
    Dim retVal() As Variant
    On Error GoTo Exit_
    
    Dim recArray As Variant
    recArray = records.GetRows
    Dim nrRows As Integer
    nrRows = UBound(recArray, 2) + 1
    recArray = TransposeDim(recArray)

    Dim fltrdArr() As Variant
    Dim MinR As Integer
    Dim MaxR As Integer
    Dim MinC As Integer
    Dim MaxC As Integer
    Dim blnY As Boolean
    MinR = LBound(recArray, 1)
    MaxR = UBound(recArray, 1)
    MinC = LBound(recArray, 2)
    MaxC = UBound(recArray, 2)
    ReDim fltrdArr(MinR To MaxR, MinC To MaxC)
    Dim rSrc As Integer
    Dim rtgt As Integer
    Dim cc   As Integer
    rtgt = MinR
    For rSrc = MinR To MaxR
        blnY = z_Contains(astrFilter, recArray(rSrc, intCol))
        If blnIncludeFilter Then
            If blnY Then
                For cc = MinC To MaxC
                    fltrdArr(rtgt, cc) = recArray(rSrc, cc)
                Next cc
                rtgt = rtgt + 1
            End If
        Else ' filter OUT
            If Not blnY Then
                For cc = MinC To MaxC
                    fltrdArr(rtgt, cc) = recArray(rSrc, cc)
                Next cc
                rtgt = rtgt + 1
            End If
        End If
    Next rSrc
    
    If showHeadings Then
        ReDim retVal(MinR To rtgt, MinC To MaxC)
        Dim col As Integer
        Dim row As Integer
        For col = MinC To MaxC
            retVal(0, col) = records.Fields(col).Name
        Next
        For row = 1 To rtgt
            For col = MinC To MaxC
                retVal(row, col) = fltrdArr(row - 1, col)
            Next col
        Next row
    Else
        retVal = fltrdArr
    End If

    
    
Exit_:
    FilteredRecordSet2Array = retVal
    
End Function


Private Function _
z_Contains(astrFilter() As String, strTgt) _
As Boolean
    Dim blnRV As Boolean
    On Error GoTo Err_
    
    Dim rr As Integer
    For rr = LBound(astrFilter) To UBound(astrFilter)
        If astrFilter(rr, 0) = strTgt Then
            blnRV = True
            Exit For
        End If
    Next rr
    
    GoTo Exit_
    
Err_:
    MsgBox "z_Contains:  " & Err.Description, vbExclamation, "Error"
    
Exit_:
    z_Contains = blnRV
    
End Function


Public Function _
rsTimeSeries2Array( _
            records As Recordset, _
            Optional showHeadings As Boolean = True, _
            Optional SkipWeekends As Boolean = False) _
As Variant
    Dim retVal() As Variant
    On Error GoTo Exit_
    
    Dim recArray As Variant
    recArray = records.GetRows
    Dim nrRows As Integer
    nrRows = UBound(recArray, 2) + 1
    recArray = TransposeDim(recArray)
    
    If SkipWeekends Then recArray = z_SkipWeekends(recArray)

    If showHeadings Then
        Dim fldCount As Integer
        fldCount = records.Fields.Count
        ReDim retVal(0 To nrRows, 0 To fldCount - 1)
        Dim col As Integer
        Dim row As Integer
        For col = 0 To fldCount - 1
            retVal(0, col) = records.Fields(col).Name
        Next
        For row = 1 To nrRows
            For col = 0 To fldCount - 1
                retVal(row, col) = recArray(row - 1, col)
            Next col
        Next row
    Else
        retVal = recArray
    End If

    
    
Exit_:
    rsTimeSeries2Array = retVal
    
End Function





'-------------------------------------------------------------------
' Private functions
'-------------------------------------------------------------------
Private Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array
    On Error GoTo Err_
    
    Dim x As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For x = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(x, Y) = v(Y, x)
        Next Y
    Next x
    
    GoTo Exit_
    
Err_:
    MsgBox Err.Description
    
Exit_:
    TransposeDim = tempArray

End Function


Private Function _
z_SkipWeekends _
           (varSrc As Variant) _
As Variant
    On Error GoTo Exit_
    
    Dim varTgt As Variant
    Dim nrRows As Integer
    Dim nrCols As Integer
    Dim nrWkends As Integer
    Dim dayOfWk  As Integer
    
    nrRows = UBound(varSrc, 1)
    nrCols = UBound(varSrc, 2)
    
    Dim row As Integer
    For row = 0 To nrRows
        If IsDate(varSrc(row, 0)) Then
            dayOfWk = Weekday(varSrc(row, 0))
            If dayOfWk = 7 Or dayOfWk = 1 Then nrWkends = nrWkends + 1
        End If
    Next row
    
    ReDim varTgt(nrRows - nrWkends, nrCols)
    Dim col As Integer
    Dim rtg As Integer
    For row = 0 To nrRows
        If IsDate(varSrc(row, 0)) Then
            dayOfWk = Weekday(varSrc(row, 0))
            If dayOfWk <> 7 And dayOfWk <> 1 Then
                For col = 0 To nrCols
                    varTgt(rtg, col) = varSrc(row, col)
                Next col
                rtg = rtg + 1
            End If
        Else
            For col = 0 To nrCols
                varTgt(rtg, col) = varSrc(row, col)
            Next col
            rtg = rtg + 1
        End If
    Next row
    
Exit_:
    z_SkipWeekends = varTgt
    
End Function




'-------------------> end of file
