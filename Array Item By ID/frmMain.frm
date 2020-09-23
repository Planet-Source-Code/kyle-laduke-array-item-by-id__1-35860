VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Array item by ID"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is just an example.  If you wanted to use this actually you should make it
'resize the array if there isnt a blank position to put your item into.  If you
'have any questions or comments post feed back or email me at: programdeveloper@hotmail.com
'
'-Kyle LaDuke

Dim varArray() As Variant

Private Function AddArrItem(strID As String, varData As Variant)
  Dim intUBound As Integer, intCurrentIndex As Integer
  
  intUBound = UBound(varArray)
  
  For intCurrentIndex = 0 To intUBound
  
    If varArray(intCurrentIndex, 0) = "" Then
      varArray(intCurrentIndex, 0) = strID
      varArray(intCurrentIndex, 1) = varData
    
      AddArrItem = True
      
      Exit For
    End If
    
  Next intCurrentIndex
  
End Function

Private Function GetArrItem(strID As String) As Variant
  Dim intUBound As Integer, intCurrentIndex As Integer
  
  intUBound = UBound(varArray)
  
  For intCurrentIndex = 0 To intUBound
    
    If varArray(intCurrentIndex, 0) = strID Then _
      GetArrItem = varArray(intCurrentIndex, 1): _
      Exit For

  Next intCurrentIndex
End Function

Private Function RemoveArrItem(strID As String)
  Dim intUBound As Integer, intCurrentIndex As Integer
  
  intUBound = UBound(varArray)
  
  For intCurrentIndex = 0 To intUBound
    If varArray(intCurrentIndex, 0) = strID Then
      varArray(intCurrentIndex, 0) = ""
      varArray(intCurrentIndex, 1) = ""
    End If
  Next
End Function

Private Sub cmdGetData_Click()
  On Error Resume Next 'just incase you input some bogus info for index of item to get
  
  Dim strID As String, strLocation() As String, varTheArray As Variant
  
  strID = InputBox("What information would you like to get?" & vbCrLf & "(dob, name, or array)", "Information to get...", "name")
  
  If strID <> "array" Then
    MsgBox GetArrItem(strID)
  Else
    strLocation() = Split(InputBox("What location of the array would you like to get?" & vbCrLf & "(0,0 through 4,4)", "Location to get...", "3,3"), ",")
    
    varTheArray = GetArrItem(strID)
    
    MsgBox varTheArray(strLocation(0), strLocation(1))
  End If

End Sub

Private Sub Form_Load()
  Dim varTestArray() As Variant, intCurrentIndexX As Integer, intCurrentIndexY As Integer
  
  ReDim varArray(9, 1) 'ten rows and two columns
  ReDim varTestArray(4, 4)
  
  For intCurrentIndexX = 0 To 4
  
    For intCurrentIndexY = 0 To 4
    
      varTestArray(intCurrentIndexX, intCurrentIndexY) = (intCurrentIndexX + 10) & ":" & (intCurrentIndexY + 10)
      
    Next intCurrentIndexY
    
  Next intCurrentIndexX
  
  AddArrItem "name", "Kyle"
  AddArrItem "dob", "12/09/84"
  AddArrItem "array", varTestArray

End Sub
