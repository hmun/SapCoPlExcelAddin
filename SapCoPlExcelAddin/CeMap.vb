Public Class CeMap

    Public aCeMapCol As Collection
    Public Sub New()
        aCeMapCol = New Collection
    End Sub

    Public Sub addCeMap(pCE As String, pNewCE As String)
        Dim aCeMapRec As CeMapRec
        Dim aKey As String

        aKey = pCE
        If aCeMapCol.Contains(aKey) Then
            aCeMapRec = aCeMapCol(aKey)
            aCeMapRec.setValues(pCE, pNewCE)
        Else
            aCeMapRec = New CeMapRec
            aCeMapRec.setValues(pCE, pNewCE)
            aCeMapCol.Add(aCeMapRec, aKey)
        End If

    End Sub

    Public Function map(pCE As String) As String
        Dim aCeMapRec As CeMapRec
        Dim aKey As String
        aKey = pCE
        map = pCE
        If aCeMapCol.Contains(aKey) Then
            aCeMapRec = aCeMapCol(aKey)
            map = aCeMapRec.NewCE.Value
        End If
    End Function

End Class
