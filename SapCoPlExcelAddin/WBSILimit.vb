Public Class WBSLimit

    Public aWBSLimitCol As Collection
    Public Sub New()
        aWBSLimitCol = New Collection
    End Sub

    Public Sub addWBSLimit(pWBS As String, pFP As Integer, pFY As Integer, pTP As Integer, pTY As Integer)
        Dim aWBSLimitRec As WBSLimitRec
        Dim aKey As String
        Dim aTP As Integer = If(pTP = 0, 16, pTP)
        Dim aTY As Integer = If(pTY = 0, 9999, pTY)

        aKey = pWBS
        If aWBSLimitCol.Contains(aKey) Then
            aWBSLimitRec = aWBSLimitCol(aKey)
            aWBSLimitRec.setValues(pWBS, pFP, pFY, aTP, aTY)
        Else
            aWBSLimitRec = New WBSLimitRec
            aWBSLimitRec.setValues(pWBS, pFP, pFY, aTP, aTY)
            aWBSLimitCol.Add(aWBSLimitRec, aKey)
        End If

    End Sub

    Public Function getWBSLimitRec(pKey As String) As WBSLimitRec
        Dim aWBSLimitRec As WBSLimitRec
        If aWBSLimitCol.Contains(pKey) Then
            aWBSLimitRec = aWBSLimitCol(pKey)
            getWBSLimitRec = aWBSLimitRec
        Else
            getWBSLimitRec = Nothing
        End If
    End Function

    Public Function checkWBSLimit(pWBS As String, pFiscy As Integer, pFiscP As Integer) As Boolean
        Dim aWBSLimitRec As WBSLimitRec = Nothing
        Dim aKey As String
        Dim aRet As Boolean = True
        aKey = pWBS
        If aWBSLimitCol.Contains(aKey) Then
            aWBSLimitRec = aWBSLimitCol(aKey)
        Else
            aKey = Left(pWBS, 7) & "*"
            If aWBSLimitCol.Contains(aKey) Then
                aWBSLimitRec = aWBSLimitCol(aKey)
            End If
        End If
        aRet = True
        If Not aWBSLimitRec Is Nothing Then
            aRet = aWBSLimitRec.checkLimit(pFiscy, pFiscP)
        End If
        checkWBSLimit = aRet
    End Function

End Class
