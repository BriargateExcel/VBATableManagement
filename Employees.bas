Attribute VB_Name = "Employees"
Option Explicit

' Built on 7/5/2020 12:07:17 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "Employees."

Private Type PrivateType
    Initialized as Boolean
    Dict as Dictionary
    Wkbk as Workbook
End Type ' PrivateType

Private This as PrivateType

' No application specific declarations found

Private Const pFirstNameColumn As Long = 1
Private Const pLastNameColumn As Long = 2
Private Const pEmployeeIDColumn As Long = 3
Private Const pHireDateColumn As Long = 4
Private Const pHeaderWidth As Long = 4

Private Const pFileName As String =  vbNullString
Private Const pWorksheetName As String = vbNullString
Private Const pExternalTableName As String = vbNullString

Public Property Get FirstNameColumn() As Long
    FirstNameColumn = pFirstNameColumn
End Property ' FirstNameColumn

Public Property Get LastNameColumn() As Long
    LastNameColumn = pLastNameColumn
End Property ' LastNameColumn

Public Property Get EmployeeIDColumn() As Long
    EmployeeIDColumn = pEmployeeIDColumn
End Property ' EmployeeIDColumn

Public Property Get HireDateColumn() As Long
    HireDateColumn = pHireDateColumn
End Property ' HireDateColumn

Public Property Get Headers() As Variant
    Headers = Array( _
        "First Name", _
        "Last Name", "Employee ID", _
        "Hire Date")
End Property ' Headers

Public Property Get Dict() As Dictionary
   Set Dict = This.Dict
End Property ' Dict

Public Property Get SpecificTable() As ListObject
    ' Table in this workbook
    Set SpecificTable = EmployeesSheet.ListObjects("EmployeesTable")
End Property ' SpecificTable

Public Property Get Initialized() As Boolean
   Initialized = This.Initialized
End Property ' Initialized

Public Sub Initialize()

    Const RoutineName As String = Module_Name & "Initialize"
    On Error GoTo ErrorHandler

    Dim  LocalTable As Employees_Table
    Set LocalTable = New Employees_Table

    Set This.Dict = New Dictionary
    If Table.TryCopyTableToDictionary(LocalTable, This.Dict, Employees.SpecificTable) Then
        This.Initialized = True
    Else
        ReportError "Error copying Employees table", "Routine", RoutineName
        This.Initialized = False
        GoTo Done
    End If

    If Not This.Wkbk is Nothing Then This.Wkbk.Close
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' EmployeesInitialize

Public Sub Reset()
    This.Initialized = False
    Set This.Dict = Nothing
End Sub ' Reset

Public Property Get HeaderWidth() As Long
    HeaderWidth = pHeaderWidth
End Property ' HeaderWidth

Public Property Get GetFirstNameFromEmployeeID(ByVal EmployeeID As String) As String

    Const RoutineName As String = Module_Name & "GetFirstNameFromEmployeeID"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then Employees.Initialize

    If CheckEmployeeIDExists(EmployeeID) Then
        GetFirstNameFromEmployeeID = This.Dict(EmployeeID).FirstName
    Else
        ReportError "Unrecognized EmployeeID", _
            "Routine", RoutineName, _
            "Employee ID", EmployeeID
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description' _
                "Employee ID", EmployeeID

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetFirstNameFromEmployeeID

Public Property Get GetLastNameFromEmployeeID(ByVal EmployeeID As String) As String

    Const RoutineName As String = Module_Name & "GetLastNameFromEmployeeID"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then Employees.Initialize

    If CheckEmployeeIDExists(EmployeeID) Then
        GetLastNameFromEmployeeID = This.Dict(EmployeeID).LastName
    Else
        ReportError "Unrecognized EmployeeID", _
            "Routine", RoutineName, _
            "Employee ID", EmployeeID
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description' _
                "Employee ID", EmployeeID

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetLastNameFromEmployeeID

Public Property Get GetHireDateFromEmployeeID(ByVal EmployeeID As String) As Date

    Const RoutineName As String = Module_Name & "GetHireDateFromEmployeeID"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then Employees.Initialize

    If CheckEmployeeIDExists(EmployeeID) Then
        GetHireDateFromEmployeeID = This.Dict(EmployeeID).HireDate
    Else
        ReportError "Unrecognized EmployeeID", _
            "Routine", RoutineName, _
            "Employee ID", EmployeeID
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description' _
                "Employee ID", EmployeeID

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetHireDateFromEmployeeID

Public Function CreateKey(ByVal Record As Employees_Table) As String

    Const RoutineName As String = Module_Name & "CreateKey"
    On Error GoTo ErrorHandler

    CreateKey = Record.EmployeeID

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CreateKey

Public Function TryCopyDictionaryToArray( _
    ByVal Dict As Dictionary, _
    ByRef Ary As Variant _
    ) As Boolean

    Const RoutineName As String = Module_Name & "TryCopyDictionaryToArray"
    On Error GoTo ErrorHandler

    TryCopyDictionaryToArray = True

    If Dict.Count = 0 Then
        ReportError "Error copying Employees_Table dictionary to array,", "Routine", RoutineName
        TryCopyDictionaryToArray = False
        GoTo Done
    End If

    ReDim Ary(1 To Dict.Count, 1 To 4)

    Dim I As Long
    I = 1

    Dim Record As Employees_Table
    Dim Entry As Variant
    For Each Entry In Dict.Keys
        Set Record = Dict.Item(Entry)

        Ary(I, pFirstNameColumn) = Record.FirstName
        Ary(I, pLastNameColumn) = Record.LastName
        Ary(I, pEmployeeIDColumn) = Record.EmployeeID
        Ary(I, pHireDateColumn) = Record.HireDate

        I = I + 1
    Next Entry

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' EmployeesTryCopyDictionaryToArray

Public Function TryCopyArrayToDictionary( _
       ByVal Ary As Variant, _
       ByRef Dict As Dictionary _
       ) As Boolean

    Const RoutineName As String = Module_Name & "TryCopyArrayToDictionary"
    On Error GoTo ErrorHandler

    TryCopyArrayToDictionary = True

    Dim I As Long

    Set Dict = New Dictionary

    Dim Key As String
    Dim Record as Employees_Table

    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then
        For I = 1 To UBound(Ary, 1)
            Set Record = New Employees_Table

            Record.FirstName = Ary(I, pFirstNameColumn)
            Record.LastName = Ary(I, pLastNameColumn)
            Record.EmployeeID = Ary(I, pEmployeeIDColumn)
            Record.HireDate = Ary(I, pHireDateColumn)

            Key = Employees.CreateKey(Record)

            If Not Dict.Exists(Key) then
                Dict.Add Key, Record
            Else
                ReportWarning "Duplicate key", "Routine", RoutineName, "Key", Key
                TryCopyArrayToDictionary = False
                GoTo Done
            End If
        Next I

    Else
        ReportError "Invalid Array", "Routine", RoutineName
    End If

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' EmployeesTryCopyArrayToDictionary

Public Function CheckEmployeeIDExists(ByVal EmployeeID As String) As Boolean _

    Const RoutineName As String = Module_Name & "CheckEmployeeIDExists"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then Employees.Initialize

    If EmployeeID = vbNullString Then
        CheckEmployeeIDExists = True
        Exit Function
    End If

    CheckEmployeeIDExists = This.Dict.Exists(EmployeeID)

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description' _
                "Employee ID", EmployeeID

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CheckEmployeeIDExists

Public Sub FormatArrayAndWorksheet( _
    ByRef Ary as Variant, _
    ByVal Table As ListObject)

    Const RoutineName As String = Module_Name & "EmployeesFormatArrayAndWorksheet"
    On Error GoTo ErrorHandler


Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' EmployeesFormatArrayAndWorksheet

' No application unique routines found

