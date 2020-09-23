VERSION 5.00
Begin VB.Form frmMultiSort 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sort on any element of a multi-dimensional array"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSort 
      Caption         =   "Sort in Descending Order"
      Height          =   240
      Left            =   6810
      TabIndex        =   7
      Top             =   3795
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   390
      Left            =   6780
      TabIndex        =   10
      Top             =   6180
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   6780
      MaxLength       =   4
      TabIndex        =   9
      Text            =   "1"
      Top             =   5820
      Width           =   2235
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Sort on 5th Member"
      Height          =   255
      Index           =   4
      Left            =   6795
      TabIndex        =   6
      Top             =   3270
      Width           =   2295
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Sort on 4th Member"
      Height          =   255
      Index           =   3
      Left            =   6795
      TabIndex        =   5
      Top             =   2940
      Width           =   2295
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Sort on 3rd Member"
      Height          =   255
      Index           =   2
      Left            =   6795
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Sort on 2nd Member"
      Height          =   255
      Index           =   1
      Left            =   6795
      TabIndex        =   3
      Top             =   2340
      Width           =   2295
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Sort on 1st Member"
      Height          =   255
      Index           =   0
      Left            =   6795
      TabIndex        =   2
      Top             =   2055
      Width           =   2295
   End
   Begin VB.ComboBox cboMembers 
      Height          =   315
      ItemData        =   "frmMultiSort.frx":0000
      Left            =   6720
      List            =   "frmMultiSort.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1515
      Width           =   2400
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      ItemData        =   "frmMultiSort.frx":0040
      Left            =   6690
      List            =   "frmMultiSort.frx":005C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   525
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sort by Above Member"
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   6720
      TabIndex        =   8
      ToolTipText     =   "Sort: Actual time shown below, extra time required to update listbox"
      Top             =   4140
      Width           =   2340
   End
   Begin VB.ListBox List2 
      Height          =   6105
      ItemData        =   "frmMultiSort.frx":008E
      Left            =   3360
      List            =   "frmMultiSort.frx":0090
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   525
      Width           =   3285
   End
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   30
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   510
      Width           =   3270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Search this Number on Sorted Member Column"
      Height          =   375
      Index           =   5
      Left            =   6780
      TabIndex        =   19
      Top             =   5400
      Width           =   2205
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Number of Array Members (Number of Dimensions)"
      Height          =   390
      Index           =   4
      Left            =   6945
      TabIndex        =   18
      Top             =   1050
      Width           =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Number of Array Records"
      Height          =   255
      Index           =   3
      Left            =   6870
      TabIndex        =   17
      Top             =   240
      Width           =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Actual Sort Time"
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   16
      Top             =   4770
      Width           =   1890
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Caption         =   "0 milliseconds"
      Height          =   240
      Left            =   6750
      TabIndex        =   15
      Top             =   5040
      Width           =   2310
   End
   Begin VB.Label Label1 
      Caption         =   "Sorted -- Click to find match on left"
      Height          =   255
      Index           =   1
      Left            =   3375
      TabIndex        =   14
      Top             =   240
      Width           =   3240
   End
   Begin VB.Label Label1 
      Caption         =   "Unsorted -- Click to find match on right"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   240
      Width           =   3240
   End
End
Attribute VB_Name = "frmMultiSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' updated 21 Nov to include a search function

Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
'
' This code is based on the excellent array sorting and searching algorithm
' module: mdArray.bas by Philippe Lord // Marton

' and also the contribution of...
' James Richardson


'******************************>>>
' The original source code for Philippe Lord's post can be found at
'   www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=24367&lngWId=1
' The original source code for James Richards's post can be found at
'   www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=31576&lngWId=1
'******************************>>>

' the following routines from the above authors are mostly untouched & as was
'   TriQuickSortLong
'   TriQuickSortLong2
'   SwapLongs
'   InsertionSortLong
'  ReverseLongArray

' James Richardson proposed a way of using Lord's routines to sort a
' two-dimensional, 2 member string array; however; I needed to tweak it
' to handle arrays of various member size and handle long values. Thought
' I'd post the code so others can use it. I have also included & very quick
' search routine to find the 1st array record matching the value you provide.
' If more than one record contain that value, then you would simply move
' from the returned Index, incrementing by 1, to find any additional matches.

' OF INTEREST: Lord's original post offers several good sorting algorithms,
' in addition to sorting arrays of strings, variants, bytes, integers, etc.
' He also includes an array reversal routine to return a sorted array in
' descending order. I've included that routine here, but strongly suggest that
' if speed is an issue, then you would want to avoid the routine & build two
' separate sort routines (1 in Ascending order & 1 in Descending order).
' The reversal routine sorts ascending first & then loops back thru the array
' to swap out the upper/lower array elements to return a desceding order
' (double the work vs sorting in the desired direction the 1st time)

'His post is worthwhile to download & put in your toolbox
'******************************>>>
' Last note: none of the declares below are required for the sorting routines
' They are used to display the results

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetDialogBaseUnits Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_FINDSTRINGEXACT As Long = &H158
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const LB_FINDSTRING As Long = &H18F
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

' IMPORTANT: The array structure these routines are looking for is....
' Array(number of elements,number of records)
' For example: A 2-dimensional array of 100 recs & 3 members would be
'   either Array(2,99) or Array(1 to 3,1 to 100) << same thing

Private testArray() As Long
Private bRedimArray As Boolean
' to test against a non-zero, low bound based array change the constant below
Const LowboundOfArray As Byte = 0

Private Sub TriQuickSortLong(ByRef iArray() As Long, ByVal iMemberID As Byte, Optional ByVal bSortDesc As Boolean)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim I       As Long
   Dim J       As Long
   Dim iTemp   As Long
   
   iLBound = LBound(iArray, 2)
   iUBound = UBound(iArray, 2)
   
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   TriQuickSortLong2 iArray, 4, iLBound, iUBound, iMemberID
   InsertionSortLong iArray, iLBound, iUBound, iMemberID
   
    If bSortDesc Then ReverseLongArray iArray()

End Sub

Private Sub TriQuickSortLong2(ByRef iArray() As Long, ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long, ByVal iMemberID As Byte)
   Dim I     As Long
   Dim J     As Long
   Dim iTemp() As Long
   Dim k As Long

   ReDim iTemp(LBound(iArray, 1) To UBound(iArray, 1))

   If (iMax - iMin) > iSplit Then
      I = (iMax + iMin) / 2
      
      If iArray(1, iMin) > iArray(1, I) Then SwapLongs iArray(), iMin, I
      If iArray(1, iMin) > iArray(1, iMax) Then SwapLongs iArray(), iMin, iMax
      If iArray(1, I) > iArray(1, iMax) Then SwapLongs iArray(), I, iMax
      
      J = iMax - 1
      SwapLongs iArray(), I, J
      
      I = iMin
      For k = LBound(iTemp) To UBound(iTemp)
        iTemp(k) = iArray(k, J)
      Next
        
      
      Do
         Do
            I = I + 1
         Loop While iArray(iMemberID, I) < iTemp(iMemberID)
         
         Do
            J = J - 1
         Loop While iArray(iMemberID, J) > iTemp(iMemberID) And J > iMin
' Note: Only logic modification I made was to add the "J > iMin" above
' In certain cases, J would fall below iMin & routine would crash
'//LaVolpe
         If J < I Then Exit Do
         
         SwapLongs iArray(), I, J
      Loop
      
      SwapLongs iArray(), I, iMax - 1
      
      TriQuickSortLong2 iArray, iSplit, iMin, J, iMemberID
      TriQuickSortLong2 iArray, iSplit, I + 1, iMax, iMemberID
   End If
End Sub

   Private Sub SwapLongs(ByRef iArray, Index1 As Long, Index2 As Long)
   Dim I As Long, J As Long
   For J = LBound(iArray, 1) To UBound(iArray, 1)
    I = iArray(J, Index1)
    iArray(J, Index1) = iArray(J, Index2)
    iArray(J, Index2) = I
   Next
   End Sub

Private Sub InsertionSortLong(ByRef iArray() As Long, ByVal iMin As Long, ByVal iMax As Long, iMemberID As Byte)
   Dim I     As Long
   Dim J     As Long
   Dim iTemp() As Long
   Dim k As Long
   ReDim iTemp(LBound(iArray, 1) To UBound(iArray, 1))
   
   For I = iMin + 1 To iMax
      For k = LBound(iTemp) To UBound(iTemp)
        iTemp(k) = iArray(k, I)
      Next
      J = I
      
      Do While J > iMin
         If iArray(iMemberID, J - 1) <= iTemp(iMemberID) Then Exit Do

         For k = LBound(iTemp) To UBound(iTemp)
            iArray(k, J) = iArray(k, J - 1)
         Next
         
         J = J - 1
      Loop
      
        For k = LBound(iTemp) To UBound(iTemp)
            iArray(k, J) = iTemp(k)
        Next
   Next I
End Sub

Private Sub ReverseLongArray(ByRef iArray() As Long)
   Dim iLBound As Long
   Dim iUBound As Long

   iLBound = LBound(iArray, 2)
   iUBound = UBound(iArray, 2)
   
   While iLBound < iUBound
      SwapLongs iArray(), iLBound, iUBound
   
      iLBound = iLBound + 1
      iUBound = iUBound - 1
   Wend
End Sub

Private Sub cboMembers_Click()
    On Error Resume Next
    bRedimArray = (UBound(testArray, 1) <> Val(cboMembers.Text))
    Dim I As Integer
    For I = 0 To cboMembers.ListIndex + 1
        optSort(I).Enabled = True
    Next
    For I = cboMembers.ListIndex + 2 To optSort.UBound
        optSort(I).Enabled = False
    Next
    If optSort(Val(optSort(0).Tag)) = True Then optSort(0) = True
End Sub

Private Sub cboSize_Click()
    On Error Resume Next
    bRedimArray = (UBound(testArray, 2) <> (Val(cboSize.Text)) + LowboundOfArray)
End Sub

Private Sub chkSort_Click()

ReverseLongArray testArray()
UpdateListBox

End Sub

Private Sub cmdSearch_Click()
If Val(Text1) < 1 Then
    MsgBox "Only number values greater than zero.", vbOKOnly
Else
    Dim Index As Long
    If optSort(Val(optSort(0).Tag)).Value = False Then Index = 0 Else Index = Val(optSort(0).Tag) + 1
    Index = SearchArray(testArray(), Index, Val(Text1), CBool(chkSort.Value))
    If Index < 0 Then
        MsgBox "The value of " & Val(Text1) & " not found in Member number " & Val(optSort(0).Tag) + 1 & _
            vbCrLf & "Or, the array isn't sorted on Member number" & Val(optSort(0).Tag) + 1, vbInformation, vbOKOnly
    Else
        List2.ListIndex = Index - LowboundOfArray
    End If
End If
End Sub

Private Sub Command1_Click(Index As Integer)
    
    If bRedimArray Then RandomizeArray
    bRedimArray = False
    
    Dim timeX As Long
    timeX = GetTickCount
    TriQuickSortLong testArray(), Val(optSort(0).Tag) + 1 + Index, CBool(chkSort.Value)
    timeX = GetTickCount - timeX
    lblTime = timeX & " milliseconds"
    UpdateListBox
End Sub

Private Sub UpdateListBox()
Dim sItem As String, X As Long, Y As Long
List2.Clear
LockWindowUpdate List2.hwnd ' speed up display
    For X = 0 + LowboundOfArray To UBound(testArray, 2)
        sItem = ""
        For Y = 1 To Val(cboMembers.Text)
            sItem = sItem & vbTab & testArray(Y, X)
        Next
        List2.AddItem testArray(0, X) & vbTab & sItem
    Next
LockWindowUpdate 0&
End Sub

Private Sub RandomizeArray()
    
    Randomize Timer
    List1.Clear
    LockWindowUpdate List1.hwnd ' speed up display
    
    Dim X As Integer, Y As Integer, sItem As String
    ReDim testArray(0 To Val(cboMembers.Text), 0 + LowboundOfArray To Val(cboSize.Text) - 1 + LowboundOfArray)
    
    For X = 0 + LowboundOfArray To UBound(testArray, 2)
        sItem = ""
        For Y = 1 To Val(cboMembers.Text)
            testArray(Y, X) = CLng(Rnd * (1000 / 5 * Y) + 1)
            sItem = sItem & vbTab & testArray(Y, X)
        Next
        testArray(0, X) = X + 1 - LowboundOfArray
        List1.AddItem testArray(0, X) & vbTab & sItem
    Next
    bRedimArray = False
    
    LockWindowUpdate 0&
End Sub

Private Sub Form_Load()
    Const LB_SETTABSTOPS As Long = &H192
    Dim lstBaseUnits As Long, I As Integer
    Dim TabStop() As Long
    ReDim TabStop(0 To 5)
        lstBaseUnits = (GetDialogBaseUnits() Mod 65536) / 2
        'set tab stops
        For I = 0 To 5
            TabStop(I) = -(((I + 1) * 5) * lstBaseUnits) - lstBaseUnits
        Next
        Call SendMessage(List1.hwnd, LB_SETTABSTOPS, UBound(TabStop) + 1, TabStop(0))
        Call SendMessage(List2.hwnd, LB_SETTABSTOPS, UBound(TabStop) + 1, TabStop(0))
    Erase TabStop
    bRedimArray = True
    cboMembers.ListIndex = 1
    cboSize.ListIndex = 4
    Call Command1_Click(-1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase testArray
End Sub

Private Sub List1_Click()
    If List1.ListIndex < 0 Then Exit Sub
    List2.ListIndex = FindListItem(List2.hwnd, True, True, List1.List(List1.ListIndex))
End Sub

Private Sub List2_Click()
    If List2.ListIndex < 0 Then Exit Sub
    List1.ListIndex = FindListItem(List1.hwnd, True, True, List2.List(List2.ListIndex))
End Sub

Private Sub optSort_Click(Index As Integer)
    If optSort(Index).Value Then optSort(0).Tag = Index
    Command1(0).Enabled = True
End Sub

Private Function FindListItem(ObjectHwnd As Long, bListBox As Boolean, bExactMatch As Boolean, sCriteria As String) As Long
    ' Function checks listbox contents for match of sCriteria if bListBox = True, otherwise
    ' checks combobox contents for match of sCriteria if bListBox = False
    Dim lMatchType As Long
    If bListBox = True Then
        If bExactMatch = False Then lMatchType = LB_FINDSTRING Else lMatchType = LB_FINDSTRINGEXACT
    Else
        If bExactMatch = False Then lMatchType = CB_FINDSTRING Else lMatchType = CB_FINDSTRINGEXACT
    End If
    FindListItem = SendMessage(ObjectHwnd, lMatchType, -1, ByVal sCriteria)
End Function

Private Function SearchArray(vArray() As Long, iMemberID As Long, lCriteria As Long, bSortedDesc As Boolean) As Long

' IMPORTANT: the array first must be sorted on the MemberID

' Very short, fast & effective routine for finding a value in a multidimensional array
' This function returns an Index of the 1st match it finds. If more than 1 match exist
' then you would increment 1 at a time from the returned Index to find additional
' matches until no more matches are found

On Error GoTo ReturnResult
' If an undimensioned array or single dimension array is passed, then we will
' error out of the routine once we try to set the UB value below

Dim UB As Long, LB As Long, MB As Long, lOffset As Long, lRtn As Long
' UB=upper bound of the array segment to search
' LB=lower bound of the array segment to search
' MB=middle point of the UB & LB values
' lOffset=the lower bound of the array
'   ... the search routine is zero-based & the lOffset is used to make up the
'       difference from a zero-based array to a passed non-zero based array

    lRtn = -1
'    LB = 0
    UB = UBound(vArray, 2) - LBound(vArray, 2) + 1
    MB = UB \ 2 + LB
    lOffset = LBound(vArray, 2)

' the logic is real simple...
' start at the middle of the array & see if the the value is found
' if not then divide the next section in half & try again... loop until done
' Tip: In large arrays, it may be better to place an initial search for the
'       1st or last record to see if they match since these are always the
'       last to be checked; especially if the match is expected in that position

'As an example, you would un-rem the lines below

'If vArray(iMemberID, LBound(vArray, 2) ) = lCriteria Then
'    SearchArray = LBound(vArray, 2)
'    Exit Function
'ElseIf vArray(iMemberID, UBound(vArray, 2)) = lCriteria Then
'    SearchArray = UBound(vArray, 2)
'    Exit Function
'End If

If bSortedDesc Then
' separate loops vs putting IF statements in the loop... faster
    
    Do
        
        If vArray(iMemberID, MB + lOffset) = lCriteria Then Exit Do
        If vArray(iMemberID, MB + lOffset) < lCriteria Then UB = MB - 1 Else LB = MB + 1
        MB = (UB - LB) \ 2 + LB
    Loop Until UB <= LB

Else

    Do
        
        If vArray(iMemberID, MB + lOffset) = lCriteria Then Exit Do
        If vArray(iMemberID, MB + lOffset) > lCriteria Then UB = MB - 1 Else LB = MB + 1
        MB = (UB - LB) \ 2 + LB
    Loop Until LB >= UB

End If

If vArray(iMemberID, MB + lOffset) = lCriteria Then

    ' the routine found a match but it may not be the very 1st match; which should be the
    ' match with the lowest Index value. We simply loop backwards until we run out
    ' of matches & return the Index of the earliest match in the array
    
    lRtn = MB - 1
    Do While lRtn > -1
        If vArray(iMemberID, lRtn + lOffset) <> lCriteria Then Exit Do
        lRtn = lRtn - 1
    Loop

    lRtn = lRtn + 1 + lOffset

End If

ReturnResult:
SearchArray = lRtn
End Function


