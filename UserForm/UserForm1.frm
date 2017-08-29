VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6195
   ClientLeft      =   48
   ClientTop       =   378
   ClientWidth     =   6912
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CheckBox1_Click()
'Select all CheckBox1
    If CheckBox1.Value = True Then
        For i = 0 To ListBox1.ListCount - 1
            ListBox1.Selected(i) = True
        Next i
    End If
    
    If CheckBox1.Value = False Then
        For i = 0 To ListBox1.ListCount - 1
            ListBox1.Selected(i) = False
        Next i
    End If
End Sub

Private Sub CheckBox2_Click()
'Select All Checkbox2

    Dim SheetArray() As String
    If ListBox2.ListCount > 0 Then 'prevent error if listcount is zero
        ReDim SheetArray(ListBox2.ListCount - 1)
    End If
    
    If CheckBox2.Value = True Then
        For i = 0 To ListBox2.ListCount - 1
            ListBox2.Selected(i) = True
            SheetArray(i) = ListBox2.List(i)
        Next i
        Sheets(SheetArray).Select
    End If
    
    If CheckBox2.Value = False Then
        For i = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(i) = True Then
                Sheets(ListBox1.List(i)).Select
                Exit For
            End If
        Next i
        For i = 0 To ListBox2.ListCount - 1
            ListBox2.Selected(i) = False
        Next i
    End If
End Sub

Private Sub CommandButton1_Click()
'Activate sheet button

    For i = 0 To ListBox1.ListCount - 1
        'If ListBox1.Selected(i) = True Then ListBox2.AddItem ListBox1.List(i)
        If ListBox1.Selected(i) = True Then
            Sheets(ListBox1.List(i)).Activate
            Exit Sub
        End If
    Next i
End Sub

Private Sub CommandButton2_Click()
'Remove Button

    Dim counter As Integer
    counter = 0
    
    For i = 0 To ListBox2.ListCount - 1
        If ListBox2.Selected(i - counter) Then
            ListBox2.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
    CheckBox2.Value = False
End Sub



Private Sub CommandButton3_Click()
'Add Button

    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then ListBox2.AddItem ListBox1.List(i)
    Next i
End Sub

Private Sub CommandButton4_Click()
'Print ActiveSheets Button

'With ActiveWorkbook.Sheets("SGB-NB-07").Tab
'        .Color = 255
'        .TintAndShade = 0
'    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        ActivePrinter:="Adobe PDF", IgnorePrintAreas:=False
End Sub

Private Sub CommandButton5_Click()
'Add Button for tabs that are yellow

    For i = 0 To ListBox2.ListCount - 1
        ListBox2.RemoveItem (0)
    Next i
    CheckBox2.Value = False

    With ListBox2
        For Each sheet In Worksheets
            If sheet.Tab.Color = 65535 Then .AddItem sheet.Name
        Next sheet
    End With
End Sub



Private Sub OptionButton1_Click()
    ListBox1.MultiSelect = 0
    ListBox2.MultiSelect = 0
End Sub

Private Sub OptionButton2_Click()
    ListBox1.MultiSelect = 1
    ListBox2.MultiSelect = 1
End Sub

Private Sub OptionButton3_Click()
    ListBox1.MultiSelect = 2
    ListBox2.MultiSelect = 2
End Sub

Private Sub OptionButton4_Click()
'Unfill sheet tab color

Dim SheetArray() As String
Dim element As Variant

    'get list of sheets into array
    If ListBox2.ListCount > 0 Then 'prevent error if listcount is zero
        ReDim SheetArray(ListBox2.ListCount - 1)
    
        For i = 0 To ListBox2.ListCount - 1
            SheetArray(i) = ListBox2.List(i)
        Next i
        For Each element In SheetArray
            ActiveWorkbook.Sheets(element).Tab.Color = xlAutomatic
        Next element
    End If
End Sub

Private Sub OptionButton5_Click()
'Yellow sheet tab color

Dim SheetArray() As String
Dim element As Variant

    'get list of sheets into array
    If ListBox2.ListCount > 0 Then 'prevent error if listcount is zero
        ReDim SheetArray(ListBox2.ListCount - 1)
    
        For i = 0 To ListBox2.ListCount - 1
            SheetArray(i) = ListBox2.List(i)
        Next i
        For Each element In SheetArray
            ActiveWorkbook.Sheets(element).Tab.Color = 65535
            '65535 is yellow
            '15773696 is skyblue
            '5287936 is army green
            '5296274 is green light
            '49407 is orange
            'is red
        Next element
    End If
End Sub

Private Sub OptionButton6_Click()
'Blue sheet tab color

Dim SheetArray() As String
Dim element As Variant

    'get list of sheets into array
    If ListBox2.ListCount > 0 Then 'prevent error if listcount is zero
        ReDim SheetArray(ListBox2.ListCount - 1)
    
        For i = 0 To ListBox2.ListCount - 1
            SheetArray(i) = ListBox2.List(i)
        Next i
        For Each element In SheetArray
            ActiveWorkbook.Sheets(element).Tab.Color = 15773696
        Next element
    End If
End Sub

Private Sub OptionButton7_Click()
'Green sheet tab color

Dim SheetArray() As String
Dim element As Variant

    'get list of sheets into array
    If ListBox2.ListCount > 0 Then 'prevent error if listcount is zero
        ReDim SheetArray(ListBox2.ListCount - 1)
    
        For i = 0 To ListBox2.ListCount - 1
            SheetArray(i) = ListBox2.List(i)
        Next i
        For Each element In SheetArray
            ActiveWorkbook.Sheets(element).Tab.Color = 5296274
        Next element
    End If
End Sub

Private Sub UserForm_Initialize()
    With ListBox1
        For Each sheet In Worksheets
            .AddItem sheet.Name
        Next sheet
'        .AddItem "Sales"
'        .AddItem "Production"
'        .AddItem "Logistics"
'        .AddItem "Human Resources"
    End With
    OptionButton3.Value = True
End Sub
