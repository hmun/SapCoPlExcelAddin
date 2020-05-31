' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/
Imports SAP.Middleware.Connector

Public Class SAPBapiTranctionCommit

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        aSapCon.getDestination(destination)
        oRfcFunction = destination.Repository.CreateFunction("BAPI_TRANSACTION_COMMIT")
    End Sub

    Public Function commit() As Integer
        sapcon.checkCon()
        Try
            oRfcFunction.Invoke(destination)
            commit = 0
            Exit Function
        Catch ex As Exception
            MsgBox("Exception in commit! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPBapiTranctionCommit")
            commit = 8
        End Try

    End Function
End Class
