Attribute VB_Name = "Table"
Option Explicit
' 2020-04-08
'   Added capability to read/write to/from an access database

Private Const Module_Name As String = "Table."

Public Function TryCopyDictionaryToTable( _
    ByVal TableType As iTable, _
    ByVal Dict As Dictionary, _
    Optional ByVal Tbl As ListObject = Nothing, _
    Optional ByVal Rng As Range = Nothing, _
    Optional ByVal TableName As String = vbNullString, _
    Optional CopyToTableRegardless As Boolean = False, _
    Optional ByVal KeepWorkbookOpen As Boolean = True _
    ) As Boolean
    ' This routine copies a dictionary to an Excel table or a database
    ' If Dict is nothing then use TableType.LocalDictionary
    '
    ' If Tbl is nothing then build a table using Rng and TableName
    '
    ' If Tbl and Rng are both Nothing then use TableType.LocalTable
    '
    ' CopyToTableRegardless = True forces copying to an Excel table
    '   regardless of whether there is an associated database
    '
    ' KeepWorkbookOpen = True avoids the workbook close in TryCopyDictionaryToExcelTable

    Const RoutineName As String = Module_Name & "CopyDictionaryToTable"
    On Error GoTo ErrorHandler
    
    If Dict.Count = 0 Then
        GoTo Done
    End If
    
    TryCopyDictionaryToTable = True
    
    If TableType.IsDatabase And Not CopyToTableRegardless Then
        Dim Ary As Variant
        
        If TableType.TryCopyDictionaryToArray(Dict, Ary) Then
        Else
            ReportError "Error copying dictionary to array", "Routine", RoutineName
            TryCopyDictionaryToTable = False
            GoTo Done
        End If
        
        If TryCopyArrayToDatabase( _
            Ary, TableType.DatabaseName, TableType.DatabaseTableName, False) _
        Then
        Else
            ReportError "Error copying dictionary to database", "Routine", RoutineName
            TryCopyDictionaryToTable = False
            GoTo Done
        End If
    Else
        If TryCopyDictionaryToExcelTable(TableType, Dict, Tbl, Rng, TableName, KeepWorkbookOpen) Then
        Else
            ReportError "Error copying dictionary to Excel Table", "Routine", RoutineName
            TryCopyDictionaryToTable = False
            GoTo Done
        End If
    End If

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyDictionaryToTable

Private Function TryCopyDictionaryToExcelTable( _
    ByVal TableType As iTable, _
    ByVal Dict As Dictionary, _
    Optional ByVal Tbl As ListObject = Nothing, _
    Optional ByVal Rng As Range = Nothing, _
    Optional ByVal TableName As String = vbNullString, _
    Optional ByVal KeepWorkbookOpen As Boolean = True _
    ) As Boolean
    
    ' This routine copies a dictionary to an Excel table
    ' If Dict is nothing then use TableType.LocalDictionary
    '
    ' If Tbl is nothing then build a table using Rng and TableName
    '
    ' If Tbl and Rng are both Nothing then use TableType.LocalTable
    '
    ' KeepWorkbookOpen = True avoids the workbook close

    Const RoutineName As String = Module_Name & "TryCopyDictionaryToExcelTable"
    On Error GoTo ErrorHandler
    
    TryCopyDictionaryToExcelTable = True
    
    Dim ThisDict As Dictionary
    If Dict Is Nothing Then
        If Not TableType.Initialized Then TableType.Initialize
        Set ThisDict = TableType.LocalDictionary
    Else
        If Dict.Count = 0 Then
            TryCopyDictionaryToExcelTable = False
            GoTo Done
        End If
        Set ThisDict = Dict
    End If

    Dim ThisTbl As ListObject
    If Tbl Is Nothing Then
        If Rng Is Nothing Then
            Set ThisTbl = TableType.LocalTable
        Else
            If TableName = vbNullString Then
                ReportError "Need to provide a table name", "Routine", RoutineName
                TryCopyDictionaryToExcelTable = False
                GoTo Done
            Else
                Set ThisTbl = Rng.Parent.ListObjects.Add(xlSrcRange, _
                    Range(Cells(1, 1), Cells(2, TableType.HeaderWidth)), , xlYes)
                ThisTbl.Name = TableName
            End If
        End If
    Else
        Set ThisTbl = Tbl
        ClearTable ThisTbl
    End If
    
    Dim ThisRng As Range
    Set ThisRng = ThisTbl.HeaderRowRange
    
    ThisRng.Resize(1, TableType.HeaderWidth) = TableType.Headers
    
    Dim Ary As Variant
    ReDim Ary(1 To ThisDict.Count, 1 To TableType.HeaderWidth)

    If TableType.TryCopyDictionaryToArray(ThisDict, Ary) Then
        ' Success; do nothing
    Else
        ReportError "Error copying dictionary to array", "Routine", RoutineName
        TryCopyDictionaryToExcelTable = False
        GoTo Done
    End If
    
    ' Format the worksheet
    TableType.FormatArrayAndWorksheet Ary, ThisTbl
    
    ' Move to DatabodyRange
    Set ThisRng = ThisRng.Offset(1, 0)
    ThisRng.Resize(UBound(Ary, 1), TableType.HeaderWidth) = Ary
    ThisRng.Resize(UBound(Ary, 1), TableType.HeaderWidth) = Ary ' Seems to be needed to get the column formatting right

    ThisRng.Parent.Cells.EntireColumn.AutoFit

    ThisRng.Parent.Activate
    ActiveWindow.FreezePanes = False

    ThisTbl.DataBodyRange(1, 1).Select
    ActiveWindow.FreezePanes = True
    
    Dim WorkbookWithTable As String
    WorkbookWithTable = ThisRng.Parent.Parent.Name
    
    If Not KeepWorkbookOpen Then
        Dim Wkbk As Workbook
        Set Wkbk = Workbooks(WorkbookWithTable)
        Wkbk.Save
        Wkbk.Close
    End If
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyDictionaryToExcelTable

Public Function TryCopyTableToDictionary( _
    ByVal TableType As iTable, _
    ByRef Dict As Dictionary, _
    ByVal Tbl As ListObject, _
    Optional ByVal ReadFromDatabase As Boolean = True _
    ) As Boolean

    ' Copies a table to a dictionary

    Const RoutineName As String = Module_Name & "TryCopyTableToDictionary"
    On Error GoTo ErrorHandler

    TryCopyTableToDictionary = True
    
    Dim Ary As Variant
    
    If TableType.IsDatabase And ReadFromDatabase Then
        If TryReadDatabaseToArray(TableType.DatabaseName, TableType.DatabaseTableName, Ary) Then
        Else
            ReportError "Error copying database to array", _
                        "Routine", RoutineName, _
                        "Table Type", TableType.LocalName
            TryCopyTableToDictionary = False
            GoTo Done
        End If
    Else
        If TryCopyExcelTableToArray(TableType, Ary, Tbl) Then
        Else
            ReportError "Error copying table to array", _
                        "Routine", RoutineName, _
                        "Table Type", TableType.LocalName
            TryCopyTableToDictionary = False
            GoTo Done
        End If
    End If

    Dim ThisDict As Dictionary
    If Dict Is Nothing Then
        Set ThisDict = TableType.LocalDictionary
    Else
        Set ThisDict = Dict
    End If

    If TableType.TryCopyArrayToDictionary(Ary, ThisDict) Then
        ' Success; do nothing
    Else
        ReportError "Error loading dictionary", "Routine", RoutineName
    End If

    Set Dict = ThisDict

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyTableToDictionary

Private Function TryCopyExcelTableToArray( _
    ByVal TableType As iTable, _
    ByRef Ary As Variant, _
    ByVal Tbl As ListObject _
    ) As Boolean
    
    ' Used to return a boolean and some other value(s)
    ' Returns True if successful

    Const RoutineName As String = Module_Name & "TryCopyExcelTableToArray"
    On Error GoTo ErrorHandler
    
    TryCopyExcelTableToArray = True
    
    On Error Resume Next
    Ary = Tbl.DataBodyRange
    If Err.Number <> 0 Then
        ReportError "The " & TableType.LocalName & " table is empty", "Routine", RoutineName
        TryCopyExcelTableToArray = False
        GoTo Done
    End If
    Err.Clear
    On Error GoTo ErrorHandler
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyExcelTableToArray

Public Function TryReadDatabaseToArray( _
    ByVal DatabaseName As String, _
    ByVal DatabaseTableName As String, _
    ByRef Ary As Variant, _
    Optional ByRef FieldNames As Variant _
    ) As Boolean
    
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-object-dao
' This example demonstrates Recordset objects and the Recordsets collection
' by opening four different types of Recordsets,
' enumerating the Recordsets collection of the current Database,
' and enumerating the Properties collection of each Recordset.

    Const RoutineName As String = Module_Name & "TryReadDatabaseToArray"
    On Error GoTo ErrorHandler
    
    TryReadDatabaseToArray = True
    
    Dim DB As Database
    Dim RS As Recordset
    
'    Set DB = OpenDatabase(GetDataFilesFolder & Application.PathSeparator & DatabaseName)
    
    Set RS = DB.OpenRecordset(DatabaseTableName, dbOpenTable)
    
    
    Dim I As Long
    
    If RS.RecordCount = 0 Then
        ' I had to do this because RS.RecordCount=0 even though there were 1,000s of records
        I = 0
        RS.MoveFirst
        Do While Not RS.EOF
            RS.MoveNext
            I = I + 1
        Loop
    Else
        I = RS.RecordCount
    End If
    RS.MoveFirst
    
    ReDim Ary(1 To I, 1 To RS.Fields.Count)
    
    I = 1
    
    Dim J As Long
    
    Do While Not RS.EOF
        For J = 0 To RS.Fields.Count - 1
            If IsNull(RS.Fields(J)) Then
                Ary(I, J + 1) = vbNullString
            Else
                Ary(I, J + 1) = RS.Fields(J)
            End If
        Next J
        
        RS.MoveNext
        I = I + 1
    Loop
    
    ReDim FieldNames(1 To RS.Fields.Count)
    
    Dim Fld As DAO.Field
    I = 1
    For Each Fld In RS.Fields
        FieldNames(I) = Fld.Name
        I = I + 1
    Next Fld
            
    RS.Close
    
    DB.Close

Done:
    Exit Function
ErrorHandler:
    TryReadDatabaseToArray = False
    
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryReadDatabaseToArray

Public Function TryCopyArrayToDatabase( _
    ByVal Ary As Variant, _
    ByVal DatabaseName As String, _
    ByVal DatabaseTableName As String, _
    Optional ByVal FieldNamesInFirstRow As Boolean = False _
    ) As Boolean

    ' Copies a dictionary to its corresponding database table
    ' Assumes the first row of the array contains the field names

    Const RoutineName As String = Module_Name & "TryCopyArrayToDatabase"
    On Error GoTo ErrorHandler

    TryCopyArrayToDatabase = True

    If UBound(Ary, 1) = 0 Then GoTo Done

    Dim DB As Database
'    Set DB = OpenDatabase(GetDataFilesFolder & Application.PathSeparator & DatabaseName)

    Dim SQLQuery As String
    SQLQuery = "DELETE " & DatabaseTableName & ".* FROM " & DatabaseTableName

    DB.Execute SQLQuery

    Dim RS As Recordset
    Set RS = DB.OpenRecordset(DatabaseTableName)
    
    Dim I As Long
    Dim J As Long
    Dim FirstRow As Long
    Dim FieldNumber As Long
    
    FirstRow = IIf(FieldNamesInFirstRow, LBound(Ary, 1) + 1, LBound(Ary, 1))
    
    For I = FirstRow To UBound(Ary, 1)
        RS.AddNew
        FieldNumber = 0
        For J = LBound(Ary, 2) To UBound(Ary, 2)
            RS.Fields(FieldNumber) = Ary(I, J)
            FieldNumber = FieldNumber + 1
        Next J
        RS.Update
    Next I

    RS.Close
    
    DB.Close

Done:
    Exit Function
ErrorHandler:
    RS.Close
    DB.Close
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyArrayToDatabase



