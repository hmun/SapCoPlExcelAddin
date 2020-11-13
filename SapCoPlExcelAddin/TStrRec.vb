Public Class TStrRec

    Public STRUCNAME As SAPCommon.TField
    Public FIELDNAME As SAPCommon.TField
    Public VALUE As SAPCommon.TField
    Public CURRENCY As SAPCommon.TField
    Public FORMAT As SAPCommon.TField

    Public Sub setValues(pSTRUCNAME As String, pFIELDNAME As String, pVALUE As String, Optional pCURRENCY As String = "", Optional pFORMAT As String = "")
        STRUCNAME = New SAPCommon.TField("STRUCNAME", pSTRUCNAME)
        FIELDNAME = New SAPCommon.TField("FIELDNAME", pFIELDNAME)
        VALUE = New SAPCommon.TField("VALUE", pVALUE)
        CURRENCY = New SAPCommon.TField("CURRENCY ", pCURRENCY)
        FORMAT = New SAPCommon.TField("FORMAT ", pFORMAT)
    End Sub

    Public Sub addValues(pSTRUCNAME As String, pFIELDNAME As String, pVALUE As String, Optional pCURRENCY As String = "", Optional pFORMAT As String = "")
        Dim aValue As SAPCommon.TField
        STRUCNAME = New SAPCommon.TField("STRUCNAME", pSTRUCNAME)
        FIELDNAME = New SAPCommon.TField("FIELDNAME", pFIELDNAME)
        aValue = New SAPCommon.TField("VALUE", pVALUE)
        aValue.Value = If(VALUE.Value <> "", CStr(CDbl(VALUE.Value) + CDbl(aValue.Value)), aValue.Value)
        VALUE = aValue
        CURRENCY = New SAPCommon.TField("CURRENCY ", pCURRENCY)
        FORMAT = New SAPCommon.TField("FORMAT ", pFORMAT)
    End Sub

    Public Sub subValues(pSTRUCNAME As String, pFIELDNAME As String, pVALUE As String, Optional pCURRENCY As String = "", Optional pFORMAT As String = "")
        Dim aValue As SAPCommon.TField
        STRUCNAME = New SAPCommon.TField("STRUCNAME", pSTRUCNAME)
        FIELDNAME = New SAPCommon.TField("FIELDNAME", pFIELDNAME)
        aValue = New SAPCommon.TField("VALUE", pVALUE)
        aValue.Value = If(VALUE.Value <> "", CStr(CDbl(VALUE.Value) - CDbl(aValue.Value)), CStr(0 - CDbl(aValue.Value)))
        VALUE = aValue
        CURRENCY = New SAPCommon.TField("CURRENCY ", pCURRENCY)
        FORMAT = New SAPCommon.TField("FORMAT ", pFORMAT)
    End Sub

    Public Sub mulValues(pSTRUCNAME As String, pFIELDNAME As String, pVALUE As String, Optional pCURRENCY As String = "", Optional pFORMAT As String = "")
        Dim aValue As SAPCommon.TField
        STRUCNAME = New SAPCommon.TField("STRUCNAME", pSTRUCNAME)
        FIELDNAME = New SAPCommon.TField("FIELDNAME", pFIELDNAME)
        aValue = New SAPCommon.TField("VALUE", pVALUE)
        aValue.Value = If(VALUE.Value <> "", CStr(CDbl(VALUE.Value) * CDbl(aValue.Value)), CStr(0))
        VALUE = aValue
        CURRENCY = New SAPCommon.TField("CURRENCY ", pCURRENCY)
        FORMAT = New SAPCommon.TField("FORMAT ", pFORMAT)
    End Sub

    Public Sub divValues(pSTRUCNAME As String, pFIELDNAME As String, pVALUE As String, Optional pCURRENCY As String = "", Optional pFORMAT As String = "")
        Dim aValue As SAPCommon.TField
        STRUCNAME = New SAPCommon.TField("STRUCNAME", pSTRUCNAME)
        FIELDNAME = New SAPCommon.TField("FIELDNAME", pFIELDNAME)
        aValue = New SAPCommon.TField("VALUE", pVALUE)
        aValue.Value = If(VALUE.Value <> "", CStr(CDbl(VALUE.Value) / CDbl(aValue.Value)), CStr(0))
        VALUE = aValue
        CURRENCY = New SAPCommon.TField("CURRENCY ", pCURRENCY)
        FORMAT = New SAPCommon.TField("FORMAT ", pFORMAT)
    End Sub

    Public Function getKey() As String
        Dim aKey As String
        aKey = STRUCNAME.Value & "-" & FIELDNAME.Value
        getKey = aKey
    End Function

    Public Function getKeyR() As String
        Dim aKey As String
        aKey = STRUCNAME.Value & "-" & FIELDNAME.Value
        getKeyR = aKey
    End Function

    Public Function toStringValue() As Object
        Dim aArray(6) As String
        aArray(0) = STRUCNAME.Value
        aArray(1) = FIELDNAME.Value
        aArray(2) = VALUE.Value
        aArray(3) = CURRENCY.Value
        aArray(4) = FORMAT.Value
        aArray(6) = formated()
        toStringValue = aArray
    End Function

    Public Function formated() As Object
        Dim aSAPFormat As New SAPFormat
        Dim aDec As Integer = 2
        If CURRENCY.Value <> "" Then
            If Left(FORMAT.Value, 1) = "D" Then
                aDec = CInt(Right(FORMAT.Value, Len(FORMAT.Value) - 1))
            End If
            If VALUE.Value <> "" Then
                formated = FormatNumber(CDbl(VALUE.Value), aDec, True, False, False)
            Else
                formated = FormatNumber(0, aDec, True, False, False)
            End If
        Else
            Select Case FORMAT.Value
                Case "DATE"
                    Try
                        formated = CDate(VALUE.Value).ToString("yyyyMMdd")
                    Catch Exc As System.Exception
                        formated = ""
                    End Try
                Case "PERIO"
                    formated = Right(VALUE.Value, 4) & Left(VALUE.Value, 3)
                Case Else
                    If Left(FORMAT.Value, 1) = "U" Then
                        formated = aSAPFormat.unpack(VALUE.Value, CInt(Right(FORMAT.Value, Len(FORMAT.Value) - 1)))
                    ElseIf Left(FORMAT.Value, 1) = "P" Then
                        formated = aSAPFormat.pspid(VALUE.Value, CInt(Right(FORMAT.Value, Len(FORMAT.Value) - 1)))
                    Else
                        formated = VALUE.Value
                    End If
            End Select

        End If
    End Function

End Class
