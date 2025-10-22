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
