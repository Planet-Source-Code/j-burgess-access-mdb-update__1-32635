VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Database"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "DataUpdate2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7650
      Top             =   1650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   750
      Width           =   5655
   End
   Begin VB.TextBox label2 
      Appearance      =   0  'Flat
      Height          =   2490
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3225
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update "
      Height          =   315
      Left            =   4575
      TabIndex        =   1
      Top             =   285
      Width           =   1080
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   7365
      TabIndex        =   0
      Top             =   3900
      Width           =   2460
   End
   Begin VB.Label status1 
      BackStyle       =   0  'Transparent
      Caption         =   "Press the ""Update"" button to begin"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   45
      TabIndex        =   4
      Top             =   5880
      Width           =   5850
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   0
      Left            =   -180
      Top             =   -165
      Width           =   6210
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   1
      Left            =   -135
      Top             =   6210
      Width           =   7365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Database
Dim rst As Recordset
Dim SQL As String
Dim num_fields As Integer
Dim num_build_fields As Integer
Dim FieldInfo(4) As Variant
Dim SourceTable As String
Dim sNewDBPathAndName As String
Dim X As Integer
Dim Y As Integer
Dim TempInt As Variant
Dim TD As Variant
Dim Junk As Variant

Dim tdSource As TableDef
Dim dbSource As Database
Dim fldSource As Field

Dim tdBuild As TableDef
Dim dbBuild As Database
Dim fldBuild As Field
Dim J As Boolean

Dim SourceData(4) As String
Dim BuildData(4) As String





Sub BuildTable(SourceTable As String)
Dim NewTableDef As TableDef
Dim MyField As Field
Dim qdf As QueryDef

Set dbBuild = OpenDatabase(CurrentBuildData, False, False, ";pwd=" & CurrentBuildPassword)
 Set NewTableDef = dbBuild.CreateTableDef(SourceTable)
  Set MyField = NewTableDef.CreateField("Temp_Field_X", dbText, "100")
  NewTableDef.Fields.Append MyField
  dbBuild.TableDefs.Append NewTableDef
 
 
   'Create a dummy QueryDef object.
    Set qdf = dbBuild.CreateQueryDef("", "Select * from " & SourceTable)
 
   ' Delete the old field.
    qdf.SQL = "ALTER TABLE " & SourceTable & " DROP COLUMN Temp_Field_X"
    qdf.Execute
  
 
 Set NewTableDef = Nothing
dbBuild.Close

label2 = label2 & "Added table " & SourceTable & vbCrLf





End Sub


Function CheckForTable(SourceTable As String) As Boolean

'check to see if this table exists in the build database

Set db = OpenDatabase(CurrentBuildData, False, False, ";pwd=" & CurrentBuildPassword)
 
 For Each TD In db.TableDefs
  Junk = TD.Name
  Junk = UCase(Junk)
  If Left(Junk, 4) <> "MSYS" Then
    
    If TD.Name = SourceTable Then
     CheckForTable = True
     Exit Function
     
    Else
     CheckForTable = False
    End If
    
    
  End If
 Next


End Function

Function CheckIfBuildFieldExists() As Boolean
Dim vTempField As String

'open the build database
Set dbBuild = OpenDatabase(CurrentBuildData, False, False, ";pwd=" & CurrentBuildPassword)
 Set tdBuild = dbBuild.TableDefs(SourceTable)
 
 'get number of fields in build table
 num_build_fields = tdBuild.Fields.Count
 
 'Run through the build fields - does this field exist?
 For Y = 0 To num_build_fields - 1
  vTempField = tdBuild.Fields(Y).Name
  If FieldInfo(0) = vTempField Then
   CheckIfBuildFieldExists = True
   Y = num_build_fields - 1
  Else
   CheckIfBuildFieldExists = False
  End If
  
 Next Y

Set tdBuild = Nothing
dbBuild.Close

End Function

Sub CreateBuildField()
On Error Resume Next



'open the build database
Set dbBuild = OpenDatabase(CurrentBuildData, False, False, ";pwd=" & CurrentBuildPassword)
 Set tdBuild = dbBuild.TableDefs(SourceTable)

 Set fldBuild = tdBuild.CreateField
   
   With fldBuild
    .Name = FieldInfo(0)
    .Type = FieldInfo(1)
    .Size = FieldInfo(2)
    .Attributes = FieldInfo(3)
    .AllowZeroLength = True
    .Required = False
   End With
   
   tdBuild.Fields.Append fldBuild
   
 Set fldBuild = Nothing
 Set tdBuild = Nothing

dbBuild.Close
Set dbBuild = Nothing

 
label2 = label2 & "Added field" & FieldInfo(0) & " to table " & SourceTable & vbCrLf



End Sub

Sub CompareTables(ListText As String)


'check to see if table exists in build database
J = CheckForTable(ListText)
 If J = False Then BuildTable (ListText)

Set dbSource = OpenDatabase(CurrentSourceData, False, False, ";pwd=" & CurrentSourcePassword)
 Set tdSource = dbSource.TableDefs(ListText)
 
 'get number of fields in source table
 num_fields = tdSource.Fields.Count
 
 'Run through the source fields
 For X = 0 To num_fields - 1
  FieldInfo(0) = tdSource.Fields(X).Name
  FieldInfo(1) = tdSource.Fields(X).Type
  FieldInfo(2) = tdSource.Fields(X).Size
  FieldInfo(3) = tdSource.Fields(X).Attributes
  SourceTable = ListText
  
  J = CheckIfBuildFieldExists
  If J = False Then
   CreateBuildField
  End If
  
  
 Next X


dbSource.Close

End Sub






Sub GetWorkList()
Dim TempList As New WorkList
Dim CancelFlag As Boolean
Dim CommonDir As String
Dim YesNoMsg As Integer


On Error GoTo OH_shit

AddSetToWorkList:

CancelFlag = False

'set the common directory

If Right(App.path, 1) = "\" Then
 CommonDir = App.path
Else
 CommonDir = App.path & "\"
End If
  
 'get the source database
 
 With CD1
  .DefaultExt = ".mdb"
  .Filter = "Access Database Files (*.mdb)|*.mdb"
  .CancelError = True
  .FilterIndex = 1
  .DialogTitle = "Source Database"
  .InitDir = CommonDir
  .FileName = ""
  .Action = 1
 End With
 
 'check to see if the user hit the cancel button
 
 If CancelFlag = False Then
   
   With TempList
    .SourceLocation = CD1.FileName
   End With
   
   'is there a password
   Form2.Show 1
        
   With TempList
    .SourcePassword = GetPassword
   End With
   
   GetPassword = ""
   
   
 Else
   'user hit cancel button... must want to quit..
   Command1.Caption = "Close"
   status1.Caption = "Cannot continue without complete Worklist"
   Exit Sub
 
 End If

 'get the dest database
 
 With CD1
  .DefaultExt = ".mdb"
  .Filter = "Access Database Files (*.mdb)|*.mdb"
  .CancelError = True
  .FilterIndex = 1
  .DialogTitle = "Destination Database"
  .InitDir = CommonDir
  .FileName = ""
  .Action = 1
 End With
 
 If CancelFlag = False Then
  
   With TempList
    .DestinationLocation = CD1.FileName
   End With
   
   'is there a password
   Form3.Show 1
   
   With TempList
    .DestinationPassword = GetPassword
   End With
   
   GetPassword = ""
   
 Else
   'user hit cancel button... must want to quit..
   Command1.Caption = "Close"
   status1.Caption = "Cannot continue without complete Worklist"
   Exit Sub
 
 End If

 'add to collection
 ColWorkList.Add TempList
 
 Set TempList = Nothing
 
 'ask the user if he would like to add another set
 YesNoMsg = MsgBox("Would you like to add another set to the worklist?", vbYesNo, "Worklist Builder")
  
  If YesNoMsg = vbYes Then
   GoTo AddSetToWorkList
  End If
  

Exit Sub
OH_shit:

If Err = 32755 Then
 CancelFlag = True
 Resume Next
Else
 MsgBox "An error has ocurred. # " & Err
End If

End Sub

Private Sub Command1_Click()
Dim xx As Integer
Dim yy As Integer


If Command1.Caption <> "Close" Then

  For yy = 1 To ColWorkList.Count
    
    CurrentSourceData = ColWorkList.Item(yy).SourceLocation
    CurrentSourcePassword = ColWorkList.Item(yy).SourcePassword
    
    CurrentBuildData = ColWorkList.Item(yy).DestinationLocation
    CurrentBuildPassword = ColWorkList.Item(yy).DestinationPassword
    
    GetTableNames
  
  
    For xx = 0 To List1.ListCount - 1
     CompareTables List1.List(xx)
    Next xx
  
  Next yy
   
  Command1.Caption = "Close"
  status1.Caption = "Completed. "


Else

 Unload Me

End If

End Sub


Private Sub Form_Load()
Set ColWorkList = New Collection

GetWorkList

End Sub

Sub Trigger_Status(iText)

status1.Caption = iText



End Sub
Sub GetTableNames()

Set db = OpenDatabase(CurrentSourceData, False, False, ";pwd=" & CurrentSourcePassword)

Text1 = Text1 & vbCrLf & "Retreiving tables from " & CurrentSourceData & vbCrLf


 'populate list with recordsets from selected .mdb
 List1.Clear
 For Each TD In db.TableDefs
     Junk = TD.Name
     Junk = UCase(Junk)
     If Left(Junk, 4) <> "MSYS" Then
         List1.AddItem TD.Name
         Text1 = Text1 & "Retreiving table " & TD.Name & vbCrLf
         
     End If
 Next




End Sub

