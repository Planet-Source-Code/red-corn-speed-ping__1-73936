Attribute VB_Name = "modListViewSort"
Option Explicit
DefLng A-Z

Private Const LVM_FIRST = &H1000
Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Private Const HDI_IMAGE = &H20
Private Const HDI_FORMAT = &H4

Private Const HDF_BITMAP_ON_RIGHT = &H1000
Private Const HDF_IMAGE = &H800
Private Const HDF_STRING = &H4000

Private Const HDM_FIRST = &H1200
Private Const HDM_SETITEM = (HDM_FIRST + 4)

Private Const HDF_LEFT As Long = 0
Private Const HDF_RIGHT As Long = 1
Private Const HDF_CENTER As Long = 2

Private Enum enumShow
    bShow = -1
    bHide = 0
End Enum

Private Type HDITEM
   mask     As Long
   cxy      As Long
   pszText  As String
   hbm      As Long
   cchTextMax As Long
   fmt      As Long
   lParam   As Long
   iImage   As Long
   iOrder   As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub ShowListViewSortIcon(list As MSComctlLib.ListView, _
    Optional vntSortColumn As Variant)
    
    Dim col As MSComctlLib.ColumnHeader
    Dim iSortColumn As Integer
    Dim lAlignment As Long
    
    If Not IsMissing(vntSortColumn) Then
        iSortColumn = vntSortColumn
        For Each col In list.ColumnHeaders  'set them all 'off'
            With col
                lAlignment = GetAlignment(col)
                ShowHeaderIcon .Index, 0, bHide, list, lAlignment
            End With
        Next
        ShowHeaderIcon iSortColumn + 1, list.SortOrder, bShow, list, lAlignment
    Else
        For Each col In list.ColumnHeaders
            With col
                lAlignment = GetAlignment(col)
                If .Index = list.SortKey + 1 Then
                    ShowHeaderIcon list.SortKey + 1, list.SortOrder, bShow, list, lAlignment
                Else
                    ShowHeaderIcon .Index, 0, bHide, list, lAlignment
                End If
            End With
        Next
    End If
    
End Sub

Private Function GetAlignment(col As MSComctlLib.ColumnHeader)
' Get the columns current alignment
    With col
        Select Case .Alignment
            Case lvwColumnRight
                GetAlignment = HDF_RIGHT
            Case lvwColumnCenter
                GetAlignment = HDF_CENTER
            Case Else
                GetAlignment = HDF_LEFT
        End Select
    End With
End Function

Private Sub ShowHeaderIcon(colNo As Long, imgIconNo As Long, bShowImage As enumShow, list As MSComctlLib.ListView, lAlignment As Long)
    Dim lHeader As Long
    Dim HD      As HDITEM
    
    'get a handle to the listview header component
    lHeader = SendMessage(list.hwnd, LVM_GETHEADER, 0, ByVal 0)
    
    'set up the structure entries
    With HD
        .mask = HDI_IMAGE Or HDI_FORMAT
        
        If bShowImage Then          'show
            .fmt = HDF_STRING Or HDF_IMAGE Or HDF_BITMAP_ON_RIGHT
            .iImage = imgIconNo
        Else
            .fmt = HDF_STRING       'just string
        End If
        .fmt = .fmt Or lAlignment   '2001/12/27 Add alignment
    End With
    
    'modify the header
    Call SendMessage(lHeader, HDM_SETITEM, colNo - 1, HD)
   
End Sub

Public Function FlipSort(iDirection As Integer) As Integer
' Â½Âà±Æ§Çªº¤è¦V
    If iDirection = lvwAscending Then
        FlipSort = lvwDescending
    Else
        FlipSort = lvwAscending
    End If
End Function


