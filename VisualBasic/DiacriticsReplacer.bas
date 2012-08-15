Attribute VB_Name = "NewMacros"
Sub ReplaceOldDiacritics()
Attribute ReplaceOldDiacritics.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro4"
'
' ReplaceOldDiacritics
'
'
    Call ReplaceString(ChrW(61648), "�")
    Call ReplaceString(ChrW(61618), "�")
    Call ReplaceString(ChrW(61599), "�")
    Call ReplaceString(ChrW(61674), "�")
    Call ReplaceString(ChrW(61603), "�")
    Call ReplaceString(ChrW(61679), "�")
    Call ReplaceString(ChrW(61613), "�")
    Call ReplaceString(ChrW(61583), "�")


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
