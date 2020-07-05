VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "Employees_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iTable

' Built on 7/5/2020 12:07:17 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Type PrivateType
    FirstName As String
    LastName As String
    EmployeeID As String
    HireDate As Date
End Type ' PrivateType

Private This as PrivateType

' No application specific declarations found

Public Property Get FirstName() as String
    FirstName = This.FirstName
End Property ' FirstName

Public Property Let FirstName(ByVal Param as String)
    This.FirstName = Param
End Property ' FirstName

Public Property Get LastName() as String
    LastName = This.LastName
End Property ' LastName

Public Property Let LastName(ByVal Param as String)
    This.LastName = Param
End Property ' LastName

Public Property Get EmployeeID() as String
    EmployeeID = This.EmployeeID
End Property ' EmployeeID

Public Property Let EmployeeID(ByVal Param as String)
    This.EmployeeID = Param
End Property ' EmployeeID

Public Property Get HireDate() as Date
    HireDate = This.HireDate
End Property ' HireDate

Public Property Let HireDate(ByVal Param as Date)
    This.HireDate = Param
End Property ' HireDate

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = Employees.Dict
End Property ' LocalDictionary

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = Employees.HeaderWidth
End Property ' HeaderWidth

Public Property Get iTable_Headers() As Variant
    iTable_Headers = Employees.Headers
End Property ' Headers

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = Employees.Initialized
End Property ' Initialized

Public Sub iTable_Initialize()
    Employees.Initialize
End Sub ' Initialize

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_Localtable = Employees.SpecificTable
End Property ' LocalTable

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "Employees_Table"
End Property ' LocalName

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = Employees.TryCopyArrayToDictionary(Ary, Dict)
End Function ' TryCopyArrayToDictionary

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = Employees.TryCopyDictionaryToArray(Dict, Ary)
End Function ' TryCopyDictionaryToArray

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary as Variant, _
    ByVal Table As ListObject)
    Employees.FormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

Public Property Get iTable_CreateKey(ByVal Record As iTable) As String
    iTable_CreateKey = Employees.CreateKey(Record)
End Property ' CreateKey

Public Property Get iTable_IsDatabase() As Boolean
    iTable_IsDatabase = False
End Property ' IsDictionary

Public Property Get iTable_DatabaseName() As String
    iTable_DatabaseName = ""
End Property ' DatabaseName

Public Property Get iTable_DatabaseTableName() As String
    iTable_DatabaseTableName = "Employees"
End Property ' DatabaseTableName

' No application unique routines found

