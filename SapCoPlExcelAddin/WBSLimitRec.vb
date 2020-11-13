Public Class WBSLimitRec

    Public WBS As SAPCommon.TField
    Public FP As SAPCommon.TField
    Public FY As SAPCommon.TField
    Public TP As SAPCommon.TField
    Public TY As SAPCommon.TField

    Public Sub setValues(pWBS As String, pFP As Integer, pFY As Integer, pTP As Integer, pTY As Integer)
        WBS = New SAPCommon.TField("WBS", pWBS)
        FP = New SAPCommon.TField("FP", pFP)
        FY = New SAPCommon.TField("FY", pFY)
        TP = New SAPCommon.TField("TP", pTP)
        TY = New SAPCommon.TField("TY", pTY)
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

    Public Function checkLimit(pFiscy As Integer, pFiscP As Integer) As Boolean
        Dim aRet As Boolean
        aRet = True
        '   check the lower limit
        If (pFiscy = CInt(FY.Value) And pFiscP < CInt(FP.Value)) Or (pFiscy < CInt(FY.Value)) Then
            aRet = False
        End If
        '   check the upper limit
        If (aRet = True) And ((pFiscy = CInt(TY.Value) And pFiscP > CInt(TP.Value)) Or (pFiscy > CInt(TY.Value))) Then
            aRet = False
        End If
        checkLimit = aRet
    End Function

    Public Function toStringValue() As Object
        Dim aArray(5) As String
        aArray(0) = WBS.Value
        aArray(1) = FP.Value
        aArray(2) = FY.Value
        aArray(3) = TP.Value
        aArray(4) = TY.Value
        toStringValue = aArray
    End Function

End Class
