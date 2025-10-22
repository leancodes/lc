Option Explicit

Public Sub ToggleTwoColumnTypesetting()
    Dim rng As Range, inTbl As Boolean
    Dim sec As Section, app As Application
    Set app = Word.Application
    Set rng = Selection.Range
    inTbl = Selection.Information(wdWithInTable)

    On Error GoTo CleanFail
    app.ScreenUpdating = False
    app.Options.Pagination = False

    ' Ensure there is at least one section
    If ActiveDocument.Sections.Count = 0 Then Exit Sub
    Set sec = rng.Sections(1)

    ' CASE A: Whole doc not yet two-columns → make current section two-columns
    If sec.PageSetup.TextColumns.Count < 2 Then
        MakeSectionTwoColumns sec
        GoTo CleanExit
    End If

    ' At this point, we are in a 2-col section.
    ' CASE B: User wants a single-column block here (for a table, figure, etc.)
    If inTbl Then
        ' Breaks cannot be inside a table: create the block around the table.
        CreateSingleColumnBlockAroundTable Selection.Tables(1)
    Else
        CreateSingleColumnBlockAtSelection rng
    End If

CleanExit:
    app.Options.Pagination = True
    app.ScreenUpdating = True
    Exit Sub

CleanFail:
    app.Options.Pagination = True
    app.ScreenUpdating = True
    MsgBox "Toggle failed: " & Err.Description, vbExclamation, "Typesetting Macro"
End Sub

'=== helpers ===
Private Sub MakeSectionTwoColumns(ByVal s As Section)
    With s.PageSetup.TextColumns
        .SetCount 2
        .EvenlySpaced = True
    End With
End Sub

Private Sub MakeSectionOneColumn(ByVal s As Section)
    With s.PageSetup.TextColumns
        .SetCount 1
        .EvenlySpaced = True
    End With
End Sub

Private Sub CreateSingleColumnBlockAtSelection(ByVal r As Range)
    Dim cur As Range: Set cur = r.Duplicate
    cur.Collapse wdCollapseStart
    cur.InsertBreak wdSectionBreakContinuous        ' open block
    cur.SetRange cur.End, cur.End                   ' move into new section
    MakeSectionOneColumn cur.Sections(1)

    ' Place a closing two-column section break right after the paragraph
    Dim afterR As Range: Set afterR = cur.Duplicate
    afterR.Collapse wdCollapseEnd
    afterR.InsertBreak wdSectionBreakContinuous
    afterR.SetRange afterR.End, afterR.End
    MakeSectionTwoColumns afterR.Sections(1)

    ' Return caret to start of single-column block
    cur.Select
End Sub

Private Sub CreateSingleColumnBlockAroundTable(ByVal t As Table)
    Dim beforeR As Range, afterR As Range
    Set beforeR = t.Range.Duplicate
    beforeR.Collapse wdCollapseStart
    beforeR.InsertBreak wdSectionBreakContinuous

    ' The table lives in the new middle section now:
    MakeSectionOneColumn beforeR.Sections(1)

    Set afterR = t.Range.Duplicate
    afterR.Collapse wdCollapseEnd
    afterR.InsertBreak wdSectionBreakContinuous
    MakeSectionTwoColumns afterR.Sections(1)

    t.Range.Select
End Sub


Option Explicit

'== Public entry point ==
Public Sub FinalizeForExport_MergeRedundantSections()
    Dim i As Long
    Dim sThis As Section, sNext As Section
    Dim merged As Long
    Application.ScreenUpdating = False
    On Error GoTo Fail

    ' Walk backwards so indices remain valid as we delete breaks.
    For i = ActiveDocument.Sections.Count - 1 To 1 Step -1
        Set sThis = ActiveDocument.Sections(i)
        Set sNext = ActiveDocument.Sections(i + 1)

        ' Only merge when PageSetup (incl. columns) is effectively identical.
        If PageSetupEquivalent(sThis, sNext) Then
            DeleteSectionBreakAfterSection sThis
            merged = merged + 1
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "Finalize complete. Section breaks removed: " & merged, vbInformation, "Finalize"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Finalize failed: " & Err.Description, vbExclamation, "Finalize"
End Sub

'== Helpers ==
Private Function PageSetupEquivalent(ByVal a As Section, ByVal b As Section) As Boolean
    Dim pa As PageSetup, pb As PageSetup
    Set pa = a.PageSetup: Set pb = b.PageSetup

    ' Core page metrics (tweak as needed for your templates)
    If pa.Orientation <> pb.Orientation Then GoTo NotEq
    If pa.TopMargin <> pb.TopMargin Then GoTo NotEq
    If pa.BottomMargin <> pb.BottomMargin Then GoTo NotEq
    If pa.LeftMargin <> pb.LeftMargin Then GoTo NotEq
    If pa.RightMargin <> pb.RightMargin Then GoTo NotEq
    If pa.HeaderDistance <> pb.HeaderDistance Then GoTo NotEq
    If pa.FooterDistance <> pb.FooterDistance Then GoTo NotEq
    If pa.PageWidth <> pb.PageWidth Then GoTo NotEq
    If pa.PageHeight <> pb.PageHeight Then GoTo NotEq

    ' Columns
    With pa.TextColumns
        If .Count <> pb.TextColumns.Count Then GoTo NotEq
        ' If both are 2+, check spacing/evenness to prevent subtle shifts
        If .Count > 1 Then
            If .EvenlySpaced <> pb.TextColumns.EvenlySpaced Then GoTo NotEq
            If .EvenlySpaced = False Then
                ' Compare width of first column when custom spacing is used
                If Abs(.Item(1).Width - pb.TextColumns.Item(1).Width) > 0.5 Then GoTo NotEq
                If Abs(.Spacing - pb.TextColumns.Spacing) > 0.5 Then GoTo NotEq
            End If
        End If
    End With

    ' Header/footer mode flags (content is assumed identical in this workflow)
    If a.HeadersFooters(wdHeaderFooterPrimary).LinkToPrevious <> _
       b.HeadersFooters(wdHeaderFooterPrimary).LinkToPrevious Then GoTo NotEq
    If a.PageSetup.DifferentFirstPageHeaderFooter <> _
       b.PageSetup.DifferentFirstPageHeaderFooter Then GoTo NotEq
    If a.PageSetup.OddAndEvenPagesHeaderFooter <> _
       b.PageSetup.OddAndEvenPagesHeaderFooter Then GoTo NotEq

    PageSetupEquivalent = True
    Exit Function
NotEq:
    PageSetupEquivalent = False
End Function

' Deletes the continuous/section break at the end of a section to merge with the next one
Private Sub DeleteSectionBreakAfterSection(ByVal s As Section)
    Dim r As Range
    Set r = s.Range
    ' The last character of a Section.Range is the section break marker.
    r.Collapse wdCollapseEnd
    r.MoveStart wdCharacter, -1
    r.Delete
End Sub

' Optional: quick stats
Public Sub ReportSectionStats()
    Dim blocks As Long
    blocks = (ActiveDocument.Sections.Count - 1) / 2
    MsgBox "Sections: " & ActiveDocument.Sections.Count & vbCrLf & _
           "Estimated 1-col blocks inserted: ~" & blocks, vbInformation, "Section Stats"
End Sub
