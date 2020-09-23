VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "xls query sample"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPhoneNumber 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtEmailAddress 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   2850
      TabIndex        =   2
      Top             =   930
      Width           =   1455
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "try searching for ""smith"", ""doe"", or ""andrews"""
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'''''''   MUST REFERENCE MSADO25.TLB  (Microsoft ActiveX Data Objects 2.5 Library)


Option Explicit


Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset




Private Sub Form_Load()
    LoadExcelFile App.Path & "\Worksheet.xls"
End Sub






Private Sub cmdSearch_Click()
Dim xlsReturn As Variant

    'reset fields
    txtFirstName.Text = ""
    txtEmailAddress.Text = ""
    txtPhoneNumber.Text = ""
    
    xlsReturn = QueryExcel(txtLastName.Text, "`FIRST_NAME`, `EMAIL`, `PHONE`", "Contacts", "`LAST_NAME`")

    'read split data
    On Error GoTo noResult:
    
    txtFirstName.Text = xlsReturn(0)
    txtEmailAddress.Text = xlsReturn(1)
    txtPhoneNumber.Text = xlsReturn(2)

noResult:
Resume Next
End Sub






Public Sub LoadExcelFile(xlsFile As String)
On Error Resume Next
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.ConnectionString = "Data Source= " & xlsFile & ";" & "Extended Properties=Excel 8.0;"
    cn.CursorLocation = adUseClient
    cn.Open
End Sub


Public Function QueryExcel(clliCode As String, xlsPullFields As String, xlsSheetName As String, xlsMatchField As String) As Variant
Dim sqlCODE As String, i As Integer, x As Integer
Dim strArray() As String

Set rs = New ADODB.Recordset

x = CountOccurrences(xlsPullFields, ",")

If InStr(xlsPullFields, "`") < 1 Or InStr(xlsMatchField, "`") < 1 Then
    QueryExcel = 0
    Exit Function
End If


ReDim strArray(x)

sqlCODE = "SELECT TOP 1 " & xlsPullFields & " FROM [" & xlsSheetName & "$] WHERE " & xlsMatchField & " = '" & clliCode & "'"

On Error GoTo err:
rs.Open sqlCODE, cn, adOpenDynamic, adLockOptimistic

i = 0
Do Until i > x
If Len(rs(i)) > 1 Then
    strArray(i) = rs(i)
Else
    strArray(i) = ""
End If
i = i + 1
Loop
    QueryExcel = strArray
Exit Function

err:
QueryExcel = 0
End Function



Public Function CountOccurrences(sourceStr As String, countChar As String) As Integer
Dim i As Long
i = 1
CountOccurrences = 0
Do Until i > Len(sourceStr)
    If Mid(sourceStr, i, Len(countChar)) = countChar Then CountOccurrences = CountOccurrences + 1
    i = i + 1
Loop
End Function



