Attribute VB_Name = "EcontTrace"
' EcontExcel v0.0.1
' (c) Igor Sheludko - https://github.com/IgorBG/Econt_Excel/
'
' Tracking Econt shipment inside Excel
'
' @class EcontTrace
' @author econt.excel@gmail.com
' @license Apache2.0 (https://opensource.org/licenses/Apache-2.0)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Includes dependency on JSON Converter for VBA (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON

Option Explicit


Private Function getToken() As String
    Dim RVar As Variant
    Dim PosToken As Long
    Const TOKEN_KEY As String = "#access_token="
    Dim WHTTP As Object
    Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5.1")
    WHTTP.Open "GET", "https://www.econt.com/ms-oauth/oauth/authorize?response_type=token&client_id=prod_main_pages", False
    WHTTP.Option(6) = False 'WinHttpRequestOption_EnableRedirects
    WHTTP.send
    
    RVar = Split(WHTTP.getResponseHeader("Location"), "&")
    PosToken = InStr(RVar(0), TOKEN_KEY)
    getToken = Mid(RVar(0), PosToken + Len(TOKEN_KEY))
End Function




Public Function traceWaybillNumber(ByVal wayBillNumber As String) As String
    Dim xmlhttp As MSXML2.XMLHTTP60
    Dim JSON As Object
    Dim reqBody As Scripting.Dictionary
    Dim Track As Scripting.Dictionary
    Dim waybills As Collection
    Dim RespDict As Scripting.Dictionary
    Dim tokenStr As String
    Dim i As Long
    Dim URL As String
    
    
    '==== SETTINGS ======
    URL = "https://www.econt.com/ms-trace/api/trace/waybill"
    Set waybills = New Collection
    Set reqBody = New Dictionary
    Set RespDict = New Dictionary
    Set xmlhttp = New MSXML2.XMLHTTP60
    '====================
    
    'Collect waybills
    waybills.Add wayBillNumber
     
    'Prepare the connection
    If tokenStr = vbNullString Then tokenStr = getToken()
    xmlhttp.Open "POST", URL, False
    xmlhttp.setRequestHeader "Authorization", "Bearer " & tokenStr
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    Call reqBody.Add("waybillNumber", waybills)
    Call reqBody.Add("source", "external")
    xmlhttp.send (JsonConverter.ConvertToJson(reqBody))
'    Debug.Print xmlhttp.responseText
    Set JSON = JsonConverter.ParseJson(xmlhttp.responseText)
    xmlhttp.abort

    i = 2
    If JSON.Exists("tracks") Then
        Set Track = JSON("tracks")(1)
        'For Each Track In JSON("tracks")
            If Track.Exists("waybillNumber") Then Call RespDict.Add(CStr(Track("waybillNumber")), Track("shortDeliveryStatus"))
        '    i = i + 1
        'Next Track
    End If
traceWaybillNumber = RespDict.Item(wayBillNumber)
End Function

Public Function econt_traceOne(ByVal wayBillNumber As String) As String
If Len(wayBillNumber) <> 13 Then Exit Function
econt_traceOne = traceWaybillNumber(wayBillNumber)
End Function


Private Sub testecont_traceOne()
Dim response As String
response = econt_traceOne("1000000000000")
End Sub
