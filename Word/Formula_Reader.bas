Attribute VB_Name = "NewMacros1"
Sub testpro()
'Copy content from two documents and put it into 1
'the final doc will be newDoc
    Dim originalDoc As Document, tempDoc As Document, newDoc As Document
    Dim myPath As String, myPath1 As String, myPath2 As String, myPath3 As String
    Dim Rng As Range

'Define the location of your files
'returnts the current path of the document that you  have open
    myPath = ActiveDocument.Path
    MsgBox ActiveDocument.Path
    myPath1 = myPath & "\1.docx"
    myPath2 = myPath & "\2.docx"
'Name of doc
    myPath3 = myPath & "\test.docx"

'creates a new doc based on the normal template
    Set newDoc = Documents.Add

'set variables to certain paths and files to copy from
    Set originalDoc = Documents.Open(myPath1)
    Set tempDoc = Documents.Open(myPath2)

'first copy
    originalDoc.Content.Copy
    Set Rng = newDoc.Content

'If you use wdCollapseEnd to collapse a range that refers
'to an entire paragraph, the range is located after the
'ending paragraph mark (the beginning of the next paragraph).
    Rng.Collapse Direction:=wdCollapseEnd
    Rng.Paste

'second copy
    tempDoc.Content.Copy
    Set Rng = newDoc.Content
    Rng.Collapse Direction:=wdCollapseEnd
    Rng.Paste

'save the new doc in this directory
    newDoc.SaveAs myPath3

    originalDoc.Close SaveChanges:=False
    tempDoc.Close SaveChanges:=False
    newDoc.Close
End Sub


Sub GetTCStats()
'Gets stats on revisions on how many words and characters
'inserted and deleted in the current document
'Source from: http://wordribbon.tips.net/T011484_Counting_Changed_Words.html
'Source for enumeration of revision types: https://msdn.microsoft.com/en-us/library/ff839110.aspx
    Dim lInsertsWords As Long
    Dim lInsertsChar As Long
    Dim lDeletesWords As Long
    Dim lDeletesChar As Long
    Dim sTemp As String
    'Dont need to define oRevisions before For Each Statement
    'Dim oRevision As Revisions
    
    lInsertsWords = 0
    lInsertsChar = 0
    lDeletesWords = 0
    lDeletesChar = 0
    
'Dim i As Long
'    With ActiveDocument
'    For i = 1 To 5
'        MsgBox "Correction " & i & "Type " & .Revisions(i).Type & vbCrLf & .Revisions(i).Range
'    Next
'    End With

'For Each oRevision In ActiveDocument.Revisions
'as Revisions In ActiveDocument.Revisions
'    MsgBox oRevision.Type
'    Exit For
'Next oRevision
  
 'Will loop through each revision and count how many inserted and deleted
 'Need to type,not copy, ActiveDocument.Revisions into code for it to work
    For Each oRevision In ActiveDocument.Revisions
        Select Case oRevision.Type
            Case 1 'wdRevisionInsert
                lInsertsChar = lInsertsChar + Len(oRevision.Range.Text)
                lInsertsWords = lInsertsWords + oRevision.Range.Words.Count
            Case 2 'wdRevisionDelete
                lDeletesChar = lDeletesChar + Len(oRevision.Range.Text)
                lDeletesWords = lDeletesWords + oRevision.Range.Words.Count
        End Select
    Next oRevision
    
    'Output string on message box explaining stats
    sTemp = "Insertions" & vbCrLf
    sTemp = sTemp & "    Words: " & lInsertsWords & vbCrLf
    sTemp = sTemp & "    Characters: " & lInsertsChar & vbCrLf
    sTemp = sTemp & "Deletions" & vbCrLf
    sTemp = sTemp & "    Words: " & lDeletesWords & vbCrLf
    sTemp = sTemp & "    Characters: " & lDeletesChar & vbCrLf
    MsgBox sTemp


End Sub

Sub SelectionInsertText()
'
' SelectionInsertText Macro
'
Dim currentSelection As Word.Selection
Set currentSelection = Application.Selection

 ' Store the user's current Overtype selection
Dim userOvertype As Boolean
userOvertype = Application.Options.Overtype

' Make sure Overtype is turned off.
        If Application.Options.Overtype Then
            Application.Options.Overtype = False
        End If

        With currentSelection

            ' Test to see if selection is an insertion point.
            If .Type = Word.WdSelectionType.wdSelectionIP Then
                .TypeText ("Inserting at insertion point. ")
                .TypeParagraph

            ElseIf .Type = Word.WdSelectionType.wdSelectionNormal Then
                ' Move to start of selection.
                If Application.Options.ReplaceSelection Then
                    .Collapse Direction:=Word.WdCollapseDirection.wdCollapseEnd
                End If
                .TypeText ("Inserting before a text block. ")
                .TypeParagraph

            Else
                ' Do nothing.
            End If
        End With

        ' Restore the user's Overtype selection
        Application.Options.Overtype = userOvertype
End Sub
Sub FormulaReader()
'
' FormulaReader Macro
'Reads text in a formula and outputs the results
'
Dim currentSelection As Word.Selection
Set currentSelection = Application.Selection

Dim objExcel As Object
Set objExcel = CreateObject("Excel.Application")

Dim x1 As String
x1 = currentSelection



'where result is inputted
With currentSelection

    If .Type = Word.WdSelectionType.wdSelectionIP Then
        MsgBox "Nothing"
    
    ElseIf .Type = Word.WdSelectionType.wdSelectionNormal Then
            'currentSelection.Font.Underline = wdUnderlineSingle
          ' Move to end of selection.
        If Application.Options.ReplaceSelection Then
          .Collapse Direction:=Word.WdCollapseDirection.wdCollapseEnd
        End If
        
            'this code takes away weird characters that results in error 2015
            'char(13) is paragraph mark
            x1 = Replace(x1, Chr(13), "")
            x1 = Replace(x1, "=", "")
            
            'Result = Excel.Application.Evaluate(x1)
            Result = Round(objExcel.Application.Evaluate(x1), 2)
            
            'in case Result has an error
            On Error GoTo InvalidValue:
            
            
            'MsgBox "This is : " & Result
            'objExcel.Quit
            Set objExcel = Nothing
           .TypeText (" " & Result)
            '.TypeParagraph
             
    
    Else
        ' Do nothing.
    End If

End With

Exit Sub

'where errors are handled
InvalidValue:
MsgBox "syntax error in: " & x1
Exit Sub

End Sub
