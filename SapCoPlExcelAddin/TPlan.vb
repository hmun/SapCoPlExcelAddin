Public Class TPlan

    Public aTPlanCol As Collection
    Private aPar As TPar

    Public Sub New(ByRef pPar As TPar)
        aTPlanCol = New Collection
        aPar = pPar
    End Sub

    Public Sub addValue(pKey As String, pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String)
        Dim aTPlanRec As TPlanRec
        If aTPlanCol.Contains(pKey) Then
            aTPlanRec = aTPlanCol(pKey)
            aTPlanRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT)
        Else
            aTPlanRec = New TPlanRec
            aTPlanRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT)
            aTPlanCol.Add(aTPlanRec, pKey)
        End If
    End Sub

    Public Sub delPlan(pKey As String)
        aTPlanCol.Remove(pKey)
    End Sub

    Public Sub checkWBSLimit(ByRef pWBSLimit As WBSLimit, pKey As String)
        Dim aTPlanRec As TPlanRec
        If aTPlanCol.Contains(pKey) Then
            aTPlanRec = aTPlanCol(pKey)
            aTPlanRec.checkWBSLimit(pWBSLimit, aPar)
        End If
    End Sub

    Public Sub mapCE(ByRef pCeMap As CeMap, pKey As String)
        Dim aTPlanRec As TPlanRec
        If aTPlanCol.Contains(pKey) Then
            aTPlanRec = aTPlanCol(pKey)
            aTPlanRec.mapCE(pCeMap, aPar)
        End If
    End Sub

    Public Sub buildValFields(pKey As String)
        Dim aTPlanRec As TPlanRec
        If aTPlanCol.Contains(pKey) Then
            aTPlanRec = aTPlanCol(pKey)
            aTPlanRec.buildValFields()
        End If
    End Sub

End Class
