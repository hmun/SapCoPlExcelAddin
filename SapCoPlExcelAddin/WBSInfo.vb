Public Class WBSInfo

    Public aWBSInfoCol As Collection
    Private aPlanningAllowedValue As String
    Private aPlanningAllowedStatus() As String

    Public Sub New(ByRef pPar As TPar)
        aWBSInfoCol = New Collection
        aPlanningAllowedValue = pPar.value("0WBS_ELEMT", "ZPSPLANEL")
        aPlanningAllowedStatus = Split(pPar.value("0WBS_ELEMT", "ZANWSTA"), ";")
    End Sub

    Public Sub addWBSInfo(pWBS As String, pPlanElement As String, pUserStatus As String)
        Dim aWBSInfoRec As WBSInfoRec
        Dim aKey As String

        aKey = pWBS
        If aWBSInfoCol.Contains(aKey) Then
            aWBSInfoRec = aWBSInfoCol(aKey)
            aWBSInfoRec.setValues(pWBS, pPlanElement, pUserStatus)
        Else
            aWBSInfoRec = New WBSInfoRec
            aWBSInfoRec.setValues(pWBS, pPlanElement, pUserStatus)
            aWBSInfoCol.Add(aWBSInfoRec, aKey)
        End If
    End Sub

    Public Function value(pSTRUCNAME As String, pFIELDNAME As String) As String
        Dim aStrRec As TStrRec
        Dim aKey As String
        aKey = pSTRUCNAME & "-" & pFIELDNAME
        If aWBSInfoCol.Contains(pSTRUCNAME & "-" & pFIELDNAME) Then
            aStrRec = aWBSInfoCol(aKey)
            value = aStrRec.VALUE.Value
        Else
            value = ""
        End If
    End Function

    Public Function isPlanningAllowed(pWBS As String) As Boolean
        Dim aWBSInfoRec As WBSInfoRec
        Dim aKey As String
        aKey = pWBS
        isPlanningAllowed = False
        If aWBSInfoCol.Contains(aKey) Then
            aWBSInfoRec = aWBSInfoCol(aKey)
            If aWBSInfoRec.PlanElement.Value = aPlanningAllowedValue Then
                isPlanningAllowed = True
            End If
        End If
    End Function

    Public Function isClosed(pWBS As String) As Boolean
        Dim aWBSInfoRec As WBSInfoRec
        Dim aKey As String
        aKey = pWBS
        ' No WBS-Info -> isClosed = True, aWBSInfoRec.UserStatus not in aPlanningAllowedStatus -> isClosed = True
        isClosed = True
        If aWBSInfoCol.Contains(aKey) Then
            aWBSInfoRec = aWBSInfoCol(aKey)
            If aPlanningAllowedStatus.Contains(aWBSInfoRec.UserStatus.Value, StringComparer.CurrentCultureIgnoreCase) Then
                isClosed = False
            End If
        End If
    End Function

End Class
