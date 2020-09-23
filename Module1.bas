Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public ColWorkList As Collection
Public CurrentSourceData As String
Public CurrentBuildData As String
Public CurrentSourcePassword As String
Public CurrentBuildPassword As String
Public GetPassword As String




