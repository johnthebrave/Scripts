Attribute VB_Name = "NewMacros"
Sub ReplaceOldDiacritics()
Attribute ReplaceOldDiacritics.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro4"
'
' ReplaceOldDiacritics
'
'
    Call ReplaceString(ChrW(61648), "î")
    Call ReplaceString(ChrW(61618), "ã")
    Call ReplaceString(ChrW(61599), "ª")
    Call ReplaceString(ChrW(61674), "º")
    Call ReplaceString(ChrW(61603), "Þ")
    Call ReplaceString(ChrW(61679), "þ")
    Call ReplaceString(ChrW(61613), "â")
    Call ReplaceString(ChrW(61583), "Î")


End Sub
Sub ReplaceString(ByVal InitialString As String, ByVal ModifiedString As String)

    With Selection.Find
        .Text = InitialString
        .Replacement.Text = ModifiedString
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
