VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dynamic Array Example"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIndexValue 
      Caption         =   "&Index Value"
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add &New"
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "&Count"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picSortedArray 
      Height          =   4095
      Left            =   1560
      ScaleHeight     =   4035
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picArray 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame fraSort 
      Caption         =   "Sorting Methods"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   4335
      Begin VB.OptionButton optSort 
         Caption         =   "&Selection Sort"
         Height          =   375
         Index           =   3
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optSort 
         Caption         =   "&Quick Sort"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optSort 
         Caption         =   "S&hell Sort"
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optSort 
         Caption         =   "&Bubble Sort"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdDeleteAll 
      Caption         =   "D&elete All"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyArray As clsArray
'============================================================================================
Private Sub cmdAdd_Click()
On Error GoTo ErrorTrap
  ' add item to array
  MyArray.Add (MyMessage("Please enter a value to be added to the array"))
  Call UpdateDisplay
  
ErrorTrap:
  Exit Sub
End Sub 'cmdAdd_Click()
'============================================================================================
Private Sub cmdCount_Click()
  ' display count of array items
  MsgBox "Array Count: " & MyArray.IndexCount + 1, vbInformation, "Array Count"
End Sub 'cmdCount_Click()
'============================================================================================
Private Sub cmdDelete_Click()
On Error GoTo ErrorTrap
  ' delete item from array
  MyArray.Delete (MyMessage("Please enter the index value to be deleted from the array"))
  If MyArray.IndexCount = -1 Then Call DisableAll
  Call UpdateDisplay
  
ErrorTrap:
  Exit Sub
End Sub 'cmdDelete_Click()
'============================================================================================
Private Sub cmdDeleteAll_Click()
  ' delete all items from array
  MyArray.DeleteAll
  Call UpdateDisplay
  Call DisableAll
End Sub 'cmdDeleteAll_Click()
'============================================================================================
Private Sub cmdFind_Click()
On Error GoTo ErrorTrap
Dim ArrayIndex As Integer
  ' find item in array
  ArrayIndex = MyArray.FindItemIndex(MyMessage("Please enter a value to find in the array"))
  If ArrayIndex <> -1 Then
    MsgBox "Found It in Array at Index: " & ArrayIndex, vbInformation, "Found It"
  Else
    MsgBox "Not Found in Array", vbInformation, "Not Found"
  End If
  
ErrorTrap:
  Exit Sub
End Sub 'cmdFind_Click()
'============================================================================================
Private Sub cmdIndexValue_Click()
On Error GoTo ErrorTrap
Dim ArrayIndex As Integer
  ' get value at array index
  ArrayIndex = MyMessage("Please enter an index to find the value stored in the array")
  If ArrayIndex > MyArray.IndexCount Then
    MsgBox "Array Index: " & ArrayIndex & vbCrLf & vbCrLf & "Not Found", vbInformation, "Not Found"
  Else
    MsgBox "The value at index " & ArrayIndex & " is: " & MyArray.Value(ArrayIndex), vbInformation, "Index Value"
  End If
  
ErrorTrap:
  Exit Sub
End Sub 'cmdIndexValue_Click()
'============================================================================================
Private Sub Form_Load()
  ' form loading...get example ready
  Set MyArray = New clsArray
  Call NoSort
  Call DisableAll
End Sub 'Form_Load()
'============================================================================================
Private Sub cmdAddNew_Click()
On Error GoTo ErrorTrap
  ' add item to array even if it exist all ready
  MyArray.AddNew (MyMessage("Please enter a value to be added to the array"))
  If MyArray.IndexCount >= 0 Then Call EnableAll
  Call UpdateDisplay

ErrorTrap:
  Exit Sub
End Sub 'cmdAddNew_Click()
'============================================================================================
Private Sub EnableAll()
Dim Counter As Integer
  ' enable all controls for use
  cmdAdd.Enabled = True
  cmdCount.Enabled = True
  cmdFind.Enabled = True
  cmdDelete.Enabled = True
  cmdDeleteAll.Enabled = True
  cmdIndexValue.Enabled = True
  For Counter = optSort.LBound To optSort.UBound
    optSort(Counter).Enabled = True
  Next Counter
  fraSort.Enabled = True
End Sub 'EnableAll()
'============================================================================================
Private Sub DisableAll()
Dim Counter As Integer
  ' disable all controls for use
  cmdAdd.Enabled = False
  cmdCount.Enabled = False
  cmdFind.Enabled = False
  cmdDelete.Enabled = False
  cmdDeleteAll.Enabled = False
  cmdIndexValue.Enabled = False
  For Counter = optSort.LBound To optSort.UBound
    optSort(Counter).Enabled = False
  Next Counter
  fraSort.Enabled = False
End Sub 'DisableAll()
'============================================================================================
Private Sub Form_Unload(Cancel As Integer)
  ' unload the form and erase the array
  MyArray.DeleteAll
End Sub 'Form_Unload(Cancel As Integer)
'============================================================================================
Private Sub optSort_Click(Index As Integer)
On Error GoTo ErrorTrap

  ' sort the array
  Select Case Index
    Case 0: 'bubble sort
      MyArray.Sort_BubbleSort
    Case 1: 'shell sort
      MyArray.Sort_ShellSort
    Case 2: 'quick sort
      MyArray.Sort_QuickSort 0, MyArray.IndexCount
    Case 3: 'selection sort
      MyArray.Sort_SelectionSort
  End Select
  Call SortedDisplay
  Call NoSort
  
ErrorTrap:
  Exit Sub
End Sub 'optSort_Click(Index As Integer)
'============================================================================================
Private Sub NoSort()
Dim Counter As Integer
  ' disable the option buttons used for sorting
  For Counter = optSort.LBound To optSort.UBound
    optSort(Counter).Value = False
  Next Counter
End Sub 'NoSort()
'============================================================================================
Private Sub UpdateDisplay()
Dim Counter As Integer
  ' display array in picture box
  picArray.Cls
  picSortedArray.Cls
  picArray.Print "Unsorted"
  For Counter = 0 To MyArray.IndexCount
      picArray.Print MyArray.Value(Counter)
  Next Counter
End Sub 'UpdateDisplay()
'============================================================================================
Private Sub SortedDisplay()
Dim Counter As Integer
  ' display sorted array in picture box
  picSortedArray.Cls
  picSortedArray.Print "Sorted"
  For Counter = 0 To MyArray.IndexCount
      picSortedArray.Print MyArray.Value(Counter)
  Next Counter
End Sub 'SortedDisplay()
'============================================================================================
Private Function MyMessage(Message As String) As Integer
  ' create input box with custom message and return integer value
  MyMessage = InputBox(Message, "Dynamic Array Example")
End Function 'MyMessage(Message As String) As Integer
'============================================================================================
