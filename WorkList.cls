VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarSourceLocation As String 'local copy
Private mvarSourcePassword As String 'local copy
Private mvarDestinationLocation As String 'local copy
Private mvarDestinationPassword As String 'local copy

Public Property Let DestinationLocation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DestinationLocation = 5
    mvarDestinationLocation = vData
End Property


Public Property Get DestinationLocation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationLocation
    DestinationLocation = mvarDestinationLocation
End Property

Public Property Let SourceLocation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SourceLocation = 5
    mvarSourceLocation = vData
End Property


Public Property Get SourceLocation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourceLocation
    SourceLocation = mvarSourceLocation
End Property

Public Property Let SourcePassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SourcePassword = 5
    mvarSourcePassword = vData
End Property


Public Property Get SourcePassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourcePassword
    SourcePassword = mvarSourcePassword
End Property



Public Property Let DestinationPassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DestinationPassword = 5
    mvarDestinationPassword = vData
End Property


Public Property Get DestinationPassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationPassword
    DestinationPassword = mvarDestinationPassword
End Property



