Public Class TSAP_Plan

    Public aTSAP_PlanCol As Collection
    Private aPar As TPar

    Private PC_Header_Fields() As String = {"CO_AREA", "FISC_YEAR", "PERIOD_FROM", "PERIOD_TO", "VERSION", "DOC_HDR_TX", "PLAN_CURRTYPE"}
    Private PC_Object_Fields() As String = {"COSTCENTER", "ACTTYPE", "CO_BUSPROC", "ORDERID", "WBS_ELEMENT"}
    Private PC_PerVal_Fields() As String = {"COST_ELEM", "RESOURCE", "TRANS_CURRE", "TRANS_CURR", "UNIT_ISO", "UNIT_OF_MEA", "FIX_VAL_PER", "VAR_VAL_PER", "FIX_QUAN_PE", "VAR_QUAN_PE", "FUND", "FUNCTION", "GRANT_NBR", "FUND_LONG", "BUDGET_PERI"}
    Private PC_Header_Key() As String = {"CO_AREA", "FISC_YEAR", "VERSION"}
    Private PC_Object_Key() As String = {"COSTCENTER", "ACTTYPE", "CO_BUSPROC", "ORDERID", "WBS_ELEMENT"}
    Private PC_PerVal_Key() As String = {"COST_ELEM", "RESOURCE""FUNCTION"}
    Private AI_Header_Fields() As String = {"CO_AREA", "FISC_YEAR", "PERIOD_FROM", "PERIOD_TO", "VERSION", "DOC_HDR_TX", "PLAN_CURRTYPE"}
    Private AI_Object_Fields() As String = {"COSTCENTER", "ACTTYPE", "CO_BUSPROC", "ORDERID", "WBS_ELEMENT"}
    Private AI_PerVal_Fields() As String = {"SEND_CCTR", "SEND_ACTIVI", "SENBUSPROC", "ORDER_CELEM", "QUANTITY_FI", "QUANTITY_VA", "SEND_FUND", "REC_FUND", "SEND_FUNCTI", "REC_FUNCTIO", "SEND_GRANT", "REC_GRANT", "UNIT_ISO", "UNIT_OF_MEA", "SEND_BUDGET", "REC_BUDGET_"}
    Private AI_Header_Key() As String = {"CO_AREA", "FISC_YEAR", "VERSION"}
    Private AI_Object_Key() As String = {"COSTCENTER", "ACTTYPE", "CO_BUSPROC", "ORDERID", "WBS_ELEMENT"}
    Private AI_PerVal_Key() As String = {"SEND_CCTR", "SEND_ACTIVITY", "SENBUSPROC", "ORDER_CELEM", "SEND_FUNCTION"}

    Public Sub New(ByRef pPar As TPar)
        aTSAP_PlanCol = New Collection
        aPar = pPar
    End Sub

    Public Sub PC_fromTPlan(pTPlan As TPlan)
        Dim aYear As String
        Dim aTPlanRec As New TPlanRec
        Dim aPCaTPlanRec As TPlanRec
        Dim aTStrRec As TStrRec
        Dim aKey As String
        For Each aTPlanRec In pTPlan.aTPlanCol
            If Not aTPlanRec.aTPlanRecCol.Contains("PERVALUE-SEND_CCTR") And Not aTPlanRec.aTPlanRecCol.Contains("PERVALUE-SEND_ACTIVITY") And
                aTPlanRec.aTPlanRecCol.Contains("PERVALUE-COST_ELEM") Then
                aYear = aTPlanRec.getYear(aPar)
                If aYear <> "" Then
                    aPCaTPlanRec = New TPlanRec
                    For Each aTStrRec In aTPlanRec.aTPlanRecCol
                        If PC_valid_Field(aTStrRec) Then
                            aPCaTPlanRec.aTPlanRecCol.Add(aTStrRec, aTStrRec.getKey())
                        End If
                    Next
                    aKey = getPCKey(aPCaTPlanRec)
                    addPlanRec(aYear, aKey, aPCaTPlanRec)
                End If
            End If
        Next
    End Sub

    Private Sub addPlanRec(pYear As String, pKey As String, pTPlanRec As TPlanRec)
        Dim aYearCol As Collection
        Dim aTPlanRec As TPlanRec
        If aTSAP_PlanCol.Contains(pYear) Then
            aYearCol = aTSAP_PlanCol(pYear)
            If aYearCol.Contains(pKey) Then
                aTPlanRec = aYearCol(pKey)
                aTPlanRec.addValues(pTPlanRec) 'add value if record already exists
            Else
                aYearCol.Add(pTPlanRec, pKey)
            End If
        Else
            aYearCol = New Collection
            aYearCol.Add(pTPlanRec, pKey)
            aTSAP_PlanCol.Add(aYearCol, pYear)
        End If

    End Sub
    Public Function PC_valid_Field(pTStrRec As TStrRec) As Boolean
        PC_valid_Field = False
        Select Case pTStrRec.STRUCNAME.Value
            Case "HEADERINFO"
                PC_valid_Field = isInArray(pTStrRec.FIELDNAME.Value, PC_Header_Fields)
            Case "COOBJECT"
                PC_valid_Field = isInArray(pTStrRec.FIELDNAME.Value, PC_Object_Fields)
            Case "PERVALUE"
                PC_valid_Field = isInArray(Left(pTStrRec.FIELDNAME.Value, 11), PC_PerVal_Fields)
        End Select
    End Function

    Public Function getPCKey(pTPlanRec As TPlanRec) As String
        Dim aTStrRec As TStrRec
        Dim aKeyField As String = ""
        Dim aRet As String = ""
        For Each aTStrRec In pTPlanRec.aTPlanRecCol
            aKeyField = ""
            Select Case aTStrRec.STRUCNAME.Value
                Case "HEADERINFO"
                    aKeyField = If(isInArray(aTStrRec.FIELDNAME.Value, PC_Header_Key), aTStrRec.VALUE.Value, "")
                Case "COOBJECT"
                    aKeyField = If(isInArray(aTStrRec.FIELDNAME.Value, PC_Object_Key), aTStrRec.VALUE.Value, "")
                Case "PERVALUE"
                    aKeyField = If(isInArray(aTStrRec.FIELDNAME.Value, PC_PerVal_Key), aTStrRec.VALUE.Value, "")
            End Select
            If aKeyField <> "" Then
                aRet = If(aRet = "", aKeyField, aRet & "-" & aKeyField)
            End If
        Next
        getPCKey = aRet
    End Function

    Public Sub AI_fromTPlan(pTPlan As TPlan)
        Dim aYear As String
        Dim aTPlanRec As New TPlanRec
        Dim aAIaTPlanRec As TPlanRec
        Dim aTStrRec As TStrRec
        Dim aKey As String
        For Each aTPlanRec In pTPlan.aTPlanCol
            If aTPlanRec.aTPlanRecCol.Contains("PERVALUE-SEND_CCTR") And aTPlanRec.aTPlanRecCol.Contains("PERVALUE-SEND_ACTIVITY") Then
                aYear = aTPlanRec.getYear(aPar)
                If aYear <> "" Then
                    aAIaTPlanRec = New TPlanRec
                    For Each aTStrRec In aTPlanRec.aTPlanRecCol
                        If AI_valid_Field(aTStrRec) Then
                            aAIaTPlanRec.aTPlanRecCol.Add(aTStrRec, aTStrRec.getKey())
                        End If
                    Next
                    aKey = getAIKey(aAIaTPlanRec)
                    addPlanRec(aYear, aKey, aAIaTPlanRec)
                End If
            End If
        Next
    End Sub

    Public Function AI_valid_Field(pTStrRec As TStrRec) As Boolean
        AI_valid_Field = False
        Select Case pTStrRec.STRUCNAME.Value
            Case "HEADERINFO"
                AI_valid_Field = isInArray(pTStrRec.FIELDNAME.Value, AI_Header_Fields)
            Case "COOBJECT"
                AI_valid_Field = isInArray(pTStrRec.FIELDNAME.Value, AI_Object_Fields)
            Case "PERVALUE"
                AI_valid_Field = isInArray(Left(pTStrRec.FIELDNAME.Value, 11), AI_PerVal_Fields)
        End Select
    End Function

    Public Function getAIKey(pTPlanRec As TPlanRec) As String
        Dim aTStrRec As TStrRec
        Dim aKeyField As String = ""
        Dim aRet As String = ""
        For Each aTStrRec In pTPlanRec.aTPlanRecCol
            aKeyField = ""
            Select Case aTStrRec.STRUCNAME.Value
                Case "HEADERINFO"
                    aKeyField = If(isInArray(aTStrRec.FIELDNAME.Value, AI_Header_Key), aTStrRec.VALUE.Value, "")
                Case "COOBJECT"
                    aKeyField = If(isInArray(aTStrRec.FIELDNAME.Value, AI_Object_Key), aTStrRec.VALUE.Value, "")
                Case "PERVALUE"
                    aKeyField = If(isInArray(aTStrRec.FIELDNAME.Value, AI_PerVal_Key), aTStrRec.VALUE.Value, "")
            End Select
            If aKeyField <> "" Then
                aRet = If(aRet = "", aKeyField, aRet & "-" & aKeyField)
            End If
        Next
        getAIKey = aRet
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

End Class