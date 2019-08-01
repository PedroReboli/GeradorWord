Imports Microsoft.Office.Interop

Module acesso
    Public Function modi(ByVal antigo As String, ByVal novo As String, ByVal tudo As Boolean) As String
        If tudo = True Then
            Form1.objDoc.Content.Find.Execute(FindText:=antigo, ReplaceWith:=novo, Replace:=Word.WdReplace.wdReplaceAll)
        Else
            Form1.objDoc.Content.Find.Execute(FindText:=antigo, ReplaceWith:=novo, Replace:=Word.WdReplace.wdReplaceOne)
        End If
        Return ""
    End Function


End Module
