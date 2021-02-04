VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const InitialSize = 1
Private ArrayElements As Integer
Private MyArray() As Integer
'============================================================================================
Private Sub Class_Initialize()
' initialize array in memory
  ' do nothing right now
End Sub 'Class_Initialize()
'============================================================================================
Private Sub Class_Terminate()
' unload array from memory
  DeleteAll
End Sub 'Class_Terminate()
'============================================================================================
Public Sub Add(ElementValue As Integer)
' add to array only if value not there yet
  If FindItemIndex(ElementValue) = -1 Then AddNew (ElementValue)
End Sub 'Add(ElementValue As Integer)
'============================================================================================
Public Function Value(ElementIndex As Integer) As Integer
On Error Resume Next
' get and return value in array
  Value = MyArray(ElementIndex)
End Function 'Value(ElementIndex As Integer) As Integer
'============================================================================================
Public Sub AddNew(ElementValue As Integer)
' add to array regardless if value is there
  ResizeArray (True)
  MyArray(ArrayElements) = ElementValue
  'ArrayElements = ArrayElements + 1
End Sub 'AddNew(ElementValue As Integer)
'============================================================================================
Public Sub Delete(ArrayIndex As Integer)
' delete an item from the array
  Dim TempArray() As Integer
  Dim ElementCounter As Integer
  Dim Counter As Integer
  
  On Error GoTo ErrorTrap
  ElementCounter = 0
  For Counter = 0 To UBound(MyArray)
    If Counter <> ArrayIndex Then
      ReDim Preserve TempArray(ElementCounter)
      TempArray(ElementCounter) = MyArray(Counter)
      ElementCounter = ElementCounter + 1
    End If
  Next Counter
  
  If UBound(MyArray) = UBound(TempArray) Then
    MsgBox "Array Index: " & ArrayIndex & vbCrLf & vbCrLf & "Not Found", _
    vbInformation, "Not Found"
    Exit Sub
  End If
  
  ReDim MyArray(UBound(TempArray))
  MyArray() = TempArray()
  Erase TempArray
  ArrayElements = UBound(MyArray)
  Exit Sub
  
ErrorTrap:
  If Err.Number = 9 Then DeleteAll
End Sub 'Delete(ArrayIndex As Integer)
'============================================================================================
Public Sub DeleteAll()
' delete all items from the array
  Erase MyArray
  ArrayElements = 0
End Sub 'DeleteAll()
'============================================================================================
Public Function FindItemIndex(ValueToFind As Integer) As Integer
' find in array and return array index
  Dim Counter As Integer
  For Counter = 0 To IndexCount '- 1
    If MyArray(Counter) = ValueToFind Then
      FindItemIndex = Counter
      Exit Function
    End If
  Next
  ' No match was found so return -1
  FindItemIndex = -1
End Function 'FindItemIndex(ValueToFind As Integer) As Integer
'============================================================================================
Public Sub Sort_BubbleSort()
' sort array using bubble sort
  Dim Counter As Integer
  Dim TempCounter As Integer
  Dim TempValue As Integer
  
  For Counter = UBound(MyArray) To LBound(MyArray) Step -1
    For TempCounter = LBound(MyArray) + 1 To Counter
      If MyArray(TempCounter - 1) > MyArray(TempCounter) Then
        TempValue = MyArray(TempCounter - 1)
        MyArray(TempCounter - 1) = MyArray(TempCounter)
        MyArray(TempCounter) = TempValue
      End If
    Next TempCounter
  Next Counter
End Sub 'Sort_BubbleSort()
'============================================================================================
Public Sub Sort_SelectionSort()
' sort array using selection sort
  Dim Counter As Integer
  Dim TempCounter As Integer
  Dim ElementIndex As Integer
  Dim TempValue As Integer
  
  For Counter = LBound(MyArray) To UBound(MyArray) - 1
    ElementIndex = Counter
    For TempCounter = Counter + 1 To UBound(MyArray)
      If MyArray(TempCounter) < MyArray(ElementIndex) Then
        ElementIndex = TempCounter
      End If
    Next TempCounter
    TempValue = MyArray(ElementIndex)
    MyArray(ElementIndex) = MyArray(Counter)
    MyArray(Counter) = TempValue
  Next Counter
End Sub 'Sort_SelectionSort()
'============================================================================================
Public Sub Sort_ShellSort()
' sort array using shell sort
  Dim Counter As Integer
  Dim ElementIndex As Integer
  Dim MinIndexValue As Integer
  Dim TempValue As Integer
  
  MinIndexValue = LBound(MyArray)
  Do Until MinIndexValue > UBound(MyArray)
    MinIndexValue = 3 * MinIndexValue + 1
  Loop
  Do Until MinIndexValue = LBound(MyArray)
    MinIndexValue = MinIndexValue / 3
    For Counter = MinIndexValue + LBound(MyArray) To UBound(MyArray)
      TempValue = MyArray(Counter)
      ElementIndex = Counter
      Do While MyArray(ElementIndex - MinIndexValue) > TempValue
        MyArray(ElementIndex) = MyArray(ElementIndex - MinIndexValue)
        ElementIndex = ElementIndex - MinIndexValue
        If ElementIndex < MinIndexValue Then Exit Do
      Loop
      MyArray(ElementIndex) = TempValue
    Next Counter
  Loop
End Sub 'Sort_ShellSort()
'============================================================================================
Public Sub Sort_QuickSort(FirstElementIndex As Integer, LastElementIndex As Integer)
' sort array using quick sort
  Dim FirstTempElement As Integer
  Dim LastTempElement As Integer
  Dim TempValueOne As Integer
  Dim TempValueTwo As Integer

  FirstTempElement = FirstElementIndex
  LastTempElement = LastElementIndex
  TempValueOne = MyArray((FirstElementIndex + LastElementIndex) / 2)
  While (FirstTempElement <= LastTempElement)
    While (MyArray(FirstTempElement) < TempValueOne And FirstTempElement < LastElementIndex)
      FirstTempElement = FirstTempElement + 1
    Wend
    While (TempValueOne < MyArray(LastTempElement) And LastTempElement > FirstElementIndex)
      LastTempElement = LastTempElement - 1
    Wend
    If (FirstTempElement <= LastTempElement) Then
      TempValueTwo = MyArray(FirstTempElement)
      MyArray(FirstTempElement) = MyArray(LastTempElement)
      MyArray(LastTempElement) = TempValueTwo
      FirstTempElement = FirstTempElement + 1
      LastTempElement = LastTempElement - 1
    End If
  Wend
  If (FirstElementIndex < LastTempElement) Then Sort_QuickSort FirstElementIndex, LastTempElement
  If (FirstTempElement < LastElementIndex) Then Sort_QuickSort FirstTempElement, LastElementIndex
End Sub 'Sort_QuickSort(FirstElementIndex As Integer, LastElementIndex As Integer)
'============================================================================================
Public Function IndexCount() As Integer
' count the number of elements in array
On Error GoTo ErrorTrap
  IndexCount = UBound(MyArray)
  Exit Function
  
ErrorTrap:
  If Err.Number = 9 Then IndexCount = -1
End Function 'IndexCount() As Integer
'============================================================================================
Public Function Display(ElementIndex As Integer) As Integer
' display the value of array element
  Display = MyArray(ElementIndex)
End Function 'Display(ElementIndex As Integer) As Integer
'============================================================================================
Public Function DisplayAll() As String
' display all the values of the array elements
  Dim Counter As Integer
  Dim Message As String
  For Counter = 0 To IndexCount '-1
    Message = Message & MyArray(Counter) & vbCrLf
  Next
  DisplayAll = Message
End Function 'DisplayAll() As String
'============================================================================================
Private Sub ResizeArray(AddArrayElement As Boolean)
' resize the array
On Error Resume Next
  If AddArrayElement = True Then
    If IndexCount = -1 Then
      ReDim Preserve MyArray(0)
    Else
      ReDim Preserve MyArray(IndexCount + InitialSize)
      ArrayElements = ArrayElements + 1
    End If
  Else
    ReDim Preserve MyArray(IndexCount - InitialSize)
    ArrayElements = ArrayElements - 1
  End If
End Sub 'ResizeArray(AddArrayElement As Boolean)
'============================================================================================
