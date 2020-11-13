Public Class WBSInfoRec

    Public WBS As SAPCommon.TField
    Public PlanElement As SAPCommon.TField
    Public UserStatus As SAPCommon.TField

    Public Sub setValues(pWBS As String, pPlanElement As String, pUserStatus As String)
        WBS = New SAPCommon.TField("WBS", pWBS)
        PlanElement = New SAPCommon.TField("PlanElement", pPlanElement)
        UserStatus = New SAPCommon.TField("UserStatus", pUserStatus)
    End Sub

    Public Function getKey() As String
        Dim aKey As String
        aKey = WBS.Value
        getKey = aKey
    End Function

    Public Function getKeyR() As String
        Dim aKey As String
        aKey = WBS.Value
        getKeyR = aKey
    End Function

    Public Function toStringValue() As Object
        Dim aArray(2) As String
        aArray(0) = WBS.Value
        aArray(1) = PlanElement.Value
        aArray(2) = UserStatus.Value
        toStringValue = aArray
    End Function

End Class
