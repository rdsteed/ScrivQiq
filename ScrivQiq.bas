Attribute VB_Name = "ScrivQiq"
Sub CutCluster()
  Dim mf As MailMergeField
  Dim f As Field
  'Copy and delete the first Qiqqa citation cluster
  If ActiveDocument.MailMerge.Fields.Count <> 0 Then
    Set mf = ActiveDocument.MailMerge.Fields(1)
    'mf.Code.Font.Size = 2 'To minify the comment
    mf.Code.Copy
    ActiveDocument.MailMerge.Fields(1).Delete
  Else
    MsgBox ("There is no citation cluster to cut")
  End If
End Sub
Sub FootnotesToClusters()
Dim en As Endnote
Dim fn As Footnote
Dim fi As MailMergeField
 'Loop to convert all footnotes/endnotes to Qiqqa citation clusters
 For Each en In ActiveDocument.Endnotes
 Code = en.Range.Text
 If Left(Code, 10) = "MERGEFIELD" Then
   Set fi = ActiveDocument.MailMerge.Fields.Add(Range:=en.Reference, Name:="[FromScrivener]")
   fi.Code.Text = Code
   fi.Locked = True
   End If
 Next
 For Each fn In ActiveDocument.Footnotes
 Code = fn.Range.Text
 If Left(Code, 10) = "MERGEFIELD" Then
   Set fi = ActiveDocument.MailMerge.Fields.Add(Range:=fn.Reference, Name:="[FromScrivener]")
   fi.Code.Text = Code
   fi.Locked = True
   End If
 Next
End Sub
