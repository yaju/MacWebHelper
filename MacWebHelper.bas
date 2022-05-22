Attribute VB_Name = "MacWebHelper"
' MacWebHelper
'
' (c)2022 yaju
'
' Twitter: @yaju
' https://github.com/yaju/MacWebHelper
'
' ==========================================================================
' MIT License
'
' Copyright (c) 2022 yaju
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==========================================================================
Option Explicit

' timeout second
Private Const web_ConnectTimeout As Long = 30
Private Const web_MaxTimeout As Long = 60

Public ConnectTimeout As Long
Public MaxTimeout As Long

Private web_CrLf As String
Private FollowRedirects As Boolean
Private Headers As Collection
Private QuerystringParams As Collection
Private baseUrl As String
Private Accept As String

Public Enum WebStatusCode
    Ok = 200
    Created = 201
    NoContent = 204
    NotModified = 304
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
    NotFound = 404
    RequestTimeout = 408
    UnsupportedMediaType = 415
    InternalServerError = 500
    BadGateway = 502
    ServiceUnavailable = 503
    GatewayTimeout = 504
End Enum

Public Enum WebMethod
    HttpGet = 0
    HttpPost = 1
    HttpPut = 2
    HttpDelete = 3
    HttpPatch = 4
    HttpHead = 5
End Enum

Public Enum WebFormat
    PlainText = 0
    Json = 1
    FormUrlEncoded = 2
    Xml = 3
    Custom = 9
End Enum

Public Enum UrlEncodingMode
    StrictUrlEncoding
    FormUrlEncoding
    QueryUrlEncoding
    CookieUrlEncoding
    PathUrlEncoding
End Enum

' ShellResult Type
Public Type shellResult
    Output As String
    ExitCode As Long
End Type

' WebResponse Type
Public Type WebResponse
    StatusCode As WebStatusCode
    StatusDescription As String
    Content As String
    Body As Variant
    Data As Object
    Headers As Collection
End Type

' WebRequest Type
Public Type WebRequest
    Headers As Dictionary
    QuerystringParams As Dictionary
    Resource As String
    FormattedResource As String
    RequestFormat As WebFormat
    method As WebMethod
    ContentType As String
    Accept As String
    ContentLength As Long
    Body As Variant
End Type

' ============================================= '
' Public Methods
' ============================================= '

' Execute the given command
Public Function ExecuteInShell(web_Command As String) As shellResult
    Dim result As String
        
    If VBA.InStr(web_Command, "driver") > 0 Then
        Initialize
        result = AppleScriptTask("shell.scpt", "ansynchandler", web_Command)
    Else
        result = AppleScriptTask("shell.scpt", "synchandler", web_Command)
    End If
    
    If VBA.InStr(result, "error") > 0 Then
        ExecuteInShell.ExitCode = -1
        ExecuteInShell.Output = result
    Else
        ExecuteInShell.ExitCode = 0
        ExecuteInShell.Output = result
    End If

End Function

' Execute the given command
Public Function Execute(ByRef Request As WebRequest) As WebResponse
    Dim web_Http As Object
    Dim web_Response As WebResponse

    Dim web_Curl As String
    Dim web_Result As shellResult

    web_Curl = PrepareCurlRequest(Request)
    web_Result = ExecuteInShell(web_Curl)
    web_Response = CreateFromCurl(Request, web_Result.Output)
    
    Set web_Http = Nothing
    Execute = web_Response

End Function

' Get JSON from the given URL
Public Function GetJson(url As String, Optional Options As Dictionary) As WebResponse
    Dim web_Request As WebRequest
        
    If Not Options Is Nothing Then
        If Options.Exists("Headers") Then
            Set web_Request.Headers = Options("Headers")
        End If
        If Options.Exists("QuerystringParams") Then
            Set web_Request.QuerystringParams = Options("QuerystringParams")
        End If
    End If
    
    web_Request.Resource = url
    web_Request.RequestFormat = WebFormat.Json
    web_Request.method = WebMethod.HttpGet
    web_Request.Accept = FormatToMediaType(web_Request.RequestFormat)
    web_Request.ContentType = FormatToMediaType(web_Request.RequestFormat)
    web_Request.ContentLength = 0
    
    GetJson = Execute(web_Request)
End Function

' Post JSON Body (`Array`, `Collection`, `Dictionary`) to the given URL
Public Function PostJson(url As String, Body As Variant, Optional Options As Dictionary) As WebResponse
    Dim web_Request As WebRequest
    
    If Not Options Is Nothing Then
        If Options.Exists("Headers") Then
            Set web_Request.Headers = Options("Headers")
        End If
        If Options.Exists("QuerystringParams") Then
            Set web_Request.QuerystringParams = Options("QuerystringParams")
        End If
    End If
    
    web_Request.Resource = url
    web_Request.RequestFormat = WebFormat.Json
    web_Request.method = WebMethod.HttpPost
    web_Request.Accept = FormatToMediaType(web_Request.RequestFormat)
    web_Request.ContentType = FormatToMediaType(web_Request.RequestFormat)
    web_Request.Body = GetBody(Body, web_Request.RequestFormat)
    web_Request.ContentLength = Len(web_Request.Body)

    PostJson = Execute(web_Request)
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub Initialize()
    ConnectTimeout = web_ConnectTimeout
    MaxTimeout = web_MaxTimeout
    FollowRedirects = True
    baseUrl = ""
    
    web_CrLf = VBA.Chr$(13)
    Set QuerystringParams = New Collection
End Sub

' `SetHeader` should be used for headers that can only be included once with a request
Private Sub SetHeader(key As String, value As Variant)
    AddOrReplaceInKeyValues Headers, key, value
End Sub

' Helper for creating `Key-Value` pair with `Dictionary`.
Private Function CreateKeyValue(key As String, value As Variant) As Dictionary
    Dim web_KeyValue As New Dictionary

    web_KeyValue("Key") = key
    web_KeyValue("Value") = value
    Set CreateKeyValue = web_KeyValue
End Function

' Helper for adding/replacing `KeyValue` in `Collection` of `KeyValue`
Private Sub AddOrReplaceInKeyValues(KeyValues As Collection, key As Variant, value As Variant)
    Dim web_KeyValue As Dictionary
    Dim web_Index As Long
    Dim web_NewKeyValue As Dictionary

    Set web_NewKeyValue = CreateKeyValue(CStr(key), value)

    web_Index = 1
    For Each web_KeyValue In KeyValues
        If web_KeyValue("Key") = key Then
            ' Replace existing
            KeyValues.Remove web_Index

            If KeyValues.Count = 0 Then
                KeyValues.Add web_NewKeyValue
            ElseIf web_Index > KeyValues.Count Then
                KeyValues.Add web_NewKeyValue, After:=web_Index - 1
            Else
                KeyValues.Add web_NewKeyValue, Before:=web_Index
            End If
            Exit Sub
        End If

        web_Index = web_Index + 1
    Next web_KeyValue

    ' Add
    KeyValues.Add web_NewKeyValue
End Sub

' Get the media-type for the given format / custom format.
Private Function FormatToMediaType(Format As WebFormat) As String
    Select Case Format
    Case WebFormat.FormUrlEncoded
        FormatToMediaType = "application/x-www-form-urlencoded;charset=UTF-8"
    Case WebFormat.Json
        FormatToMediaType = "application/json"
    Case WebFormat.Xml
        FormatToMediaType = "application/xml"
    Case Else
        FormatToMediaType = "text/plain"
    End Select
End Function

' Get the method name for the given `WebMethod`
Private Function MethodToName(method As WebMethod) As String
    Select Case method
    Case WebMethod.HttpDelete
        MethodToName = "DELETE"
    Case WebMethod.HttpPut
        MethodToName = "PUT"
    Case WebMethod.HttpPatch
        MethodToName = "PATCH"
    Case WebMethod.HttpPost
        MethodToName = "POST"
    Case WebMethod.HttpGet
        MethodToName = "GET"
    Case WebMethod.HttpHead
        MethodToName = "HEAD"
    End Select
End Function

Private Function GetBody(Body As Variant, Format As WebFormat)
    If Not VBA.IsEmpty(Body) Then
        If VBA.VarType(Body) = vbString Then
            GetBody = Body
        Else
            ' Convert body and cache
            GetBody = ConvertToFormat(Body, Format)
        End If
    End If
End Function

' Helper for converting value to given `WebFormat`.
Private Function ConvertToFormat(Obj As Variant, Format As WebFormat) As Variant
    On Error GoTo web_ErrorHandling

    Select Case Format
    Case WebFormat.Json
        ConvertToFormat = JsonConverter.ConvertToJson(Obj)
    Case WebFormat.FormUrlEncoded
        ConvertToFormat = ConvertToUrlEncoded(Obj)
    Case Else
        If VBA.VarType(Obj) = vbString Then
            ' Plain text
            ConvertToFormat = Obj
        End If
    End Select
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred during conversion" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    Err.Raise 11001, "WebHelpers.ConvertToFormat", web_ErrorDescription
End Function

' Convert `Dictionary`/`Collection` to Url-Encoded string.
Public Function ConvertToUrlEncoded(Obj As Variant, Optional EncodingMode As UrlEncodingMode = UrlEncodingMode.FormUrlEncoding) As String
    Dim web_Encoded As String

    If TypeOf Obj Is Collection Then
        Dim web_KeyValue As Dictionary

        For Each web_KeyValue In Obj
            If VBA.Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_KeyValue("Key"), web_KeyValue("Value"), EncodingMode)
        Next web_KeyValue
    Else
        Dim web_Key As Variant

        For Each web_Key In Obj.Keys()
            If Len(web_Encoded) > 0 Then: web_Encoded = web_Encoded & "&"
            web_Encoded = web_Encoded & web_GetUrlEncodedKeyValue(web_Key, Obj(web_Key), EncodingMode)
        Next web_Key
    End If

    ConvertToUrlEncoded = web_Encoded
End Function

' Get full url by joining given `WebRequest.FormattedResource` and `BaseUrl`.
Private Function GetFullUrl(ByRef Request As WebRequest) As String
    GetFullUrl = JoinUrl(baseUrl, FormattedResource(Request))
End Function

' Join Url with /
Private Function JoinUrl(LeftSide As String, RightSide As String) As String
    If Left(RightSide, 1) = "/" Then
        RightSide = Right(RightSide, Len(RightSide) - 1)
    End If
    If Right(LeftSide, 1) = "/" Then
        LeftSide = Left(LeftSide, Len(LeftSide) - 1)
    End If

    If LeftSide <> "" And RightSide <> "" Then
        JoinUrl = LeftSide & "/" & RightSide
    Else
        JoinUrl = LeftSide & RightSide
    End If
End Function

' Get `Resource` with Url Segments replaced and Querystring added.
Private Function FormattedResource(ByRef Request As WebRequest) As String
    Dim web_Segment As Variant
    Dim web_Encoding As UrlEncodingMode

    FormattedResource = Request.Resource

    ' Add querystring
    If QuerystringParams.Count > 0 Then
        If VBA.InStr(FormattedResource, "?") <= 0 Then
            FormattedResource = FormattedResource & "?"
        Else
            FormattedResource = FormattedResource & "&"
        End If

        ' For querystrings, W3C defines form-urlencoded as the required encoding,
        ' but the treatment of space -> "+" (rather than "%20") can cause issues
        '
        ' If the request format is explicitly form-urlencoded, use FormUrlEncoding (space -> "+")
        ' otherwise, use subset of RFC 3986 and form-urlencoded that should work for both cases (space -> "%20")
        If Request.RequestFormat = WebFormat.FormUrlEncoded Then
            web_Encoding = UrlEncodingMode.FormUrlEncoding
        Else
            web_Encoding = UrlEncodingMode.QueryUrlEncoding
        End If
        FormattedResource = FormattedResource & ConvertToUrlEncoded(QuerystringParams, EncodingMode:=web_Encoding)
    End If
End Function

' Prepare cURL request for given WebRequest
Private Function PrepareCurlRequest(ByRef Request As WebRequest) As String
    Dim web_Curl As String
    Dim web_KeyValue As Dictionary
    Dim web_CookieString As String

    On Error GoTo web_ErrorHandling

    web_Curl = "curl -i"

    Set Headers = New Collection
    SetHeader "Accept", Request.Accept
    If Request.method <> WebMethod.HttpGet Then
        SetHeader "Content-Type", Request.ContentType
        If Request.ContentLength <> 0 Then
            SetHeader "Content-Length", VBA.CStr(Request.ContentLength)
        End If
    End If

    ' Set timeouts
    ' (max time = resolve + sent + receive)
    web_Curl = web_Curl & " --connect-timeout " & ConnectTimeout
    web_Curl = web_Curl & " --max-time " & MaxTimeout

    ' Setup redirects
    If FollowRedirects Then
        web_Curl = web_Curl & " --location"
    End If

    ' Set headers and cookies
    For Each web_KeyValue In Headers
        web_Curl = web_Curl & " -H '" & web_KeyValue("Key") & ": " & web_KeyValue("Value") & "'"
    Next web_KeyValue

    ' Add method, data, and url
    web_Curl = web_Curl & " -X " & MethodToName(Request.method)
    web_Curl = web_Curl & " -d '" & Request.Body & "'"
    web_Curl = web_Curl & " '" & GetFullUrl(Request) & "'"

    PrepareCurlRequest = web_Curl
    Exit Function

web_ErrorHandling:

    Err.Raise 11013 + vbObjectError, "PrepareCURLRequest", _
        "An error occurred while preparing cURL request" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description
End Function

' Convert string to ANSI bytes
Private Function StringToAnsiBytes(web_Text As String) As Byte()
    Dim web_Bytes() As Byte
    Dim web_AnsiBytes() As Byte
    Dim web_ByteIndex As Long
    Dim web_AnsiIndex As Long

    If VBA.Len(web_Text) > 0 Then
        ' Take first byte from unicode bytes
        ' VBA.Int is used for floor instead of round
        web_Bytes = web_Text
        ReDim web_AnsiBytes(VBA.Int(UBound(web_Bytes) / 2))

        web_AnsiIndex = LBound(web_Bytes)
        For web_ByteIndex = LBound(web_Bytes) To UBound(web_Bytes) Step 2
            web_AnsiBytes(web_AnsiIndex) = web_Bytes(web_ByteIndex)
            web_AnsiIndex = web_AnsiIndex + 1
        Next web_ByteIndex
    End If

    StringToAnsiBytes = web_AnsiBytes
End Function

' Create response from cURL
Private Function CreateFromCurl(ByRef Request As WebRequest, result As String) As WebResponse
    On Error GoTo web_ErrorHandling

    Dim web_Lines() As String
    Dim web_Response As WebResponse

    web_Lines = VBA.Split(result, web_CrLf)
    If UBound(web_Lines) = 0 Then
        web_Response.Body = result
    Else
        web_Response.StatusCode = web_ExtractStatusFromCurlResponse(web_Lines)
        web_Response.StatusDescription = web_ExtractStatusTextFromCurlResponse(web_Lines)
        web_Response.Content = web_ExtractResponseTextFromCurlResponse(web_Lines)
        web_Response.Body = StringToAnsiBytes(web_Response.Content)
        Set Headers = ExtractHeaders(web_ExtractHeadersFromCurlResponse(web_Lines))
    End If
    
    CreateFromCurl = web_Response
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while creating response from cURL" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    Err.Raise 11031 + vbObjectError, "WebResponse.CreateFromCurl", web_ErrorDescription
End Function

' Extract headers from response headers
Public Function ExtractHeaders(ResponseHeaders As String) As Collection
    On Error GoTo web_ErrorHandling

    Dim web_Lines As Variant
    Dim web_i As Integer
    Dim web_Headers As New Collection
    Dim web_Header As Dictionary
    Dim web_ColonPosition As Long
    Dim web_Multiline As Boolean

    web_Lines = VBA.Split(ResponseHeaders, web_CrLf)

    For web_i = LBound(web_Lines) To (UBound(web_Lines) + 1)
        If web_i > UBound(web_Lines) Then
            web_Headers.Add web_Header
        ElseIf web_Lines(web_i) <> "" Then
            web_ColonPosition = VBA.InStr(1, web_Lines(web_i), ":")
            If web_ColonPosition = 0 And Not web_Header Is Nothing Then
                ' Assume part of multi-line header
                web_Multiline = True
            ElseIf web_Multiline Then
                ' Close out multi-line string
                web_Multiline = False
                web_Headers.Add web_Header
            ElseIf Not web_Header Is Nothing Then
                ' Add previous header
                web_Headers.Add web_Header
            End If

            If Not web_Multiline Then
                Set web_Header = CreateKeyValue( _
                    key:=VBA.Trim(VBA.Mid$(web_Lines(web_i), 1, web_ColonPosition - 1)), _
                    value:=VBA.Trim(VBA.Mid$(web_Lines(web_i), web_ColonPosition + 1, VBA.Len(web_Lines(web_i)))) _
                )
            Else
                web_Header("Value") = web_Header("Value") & web_CrLf & web_Lines(web_i)
            End If
        End If
    Next web_i

    Set ExtractHeaders = web_Headers
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while extracting headers" & vbNewLine & _
        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.Description

    Err.Raise 11032 + vbObjectError, "CreateFromCurl", web_ErrorDescription
End Function

Private Function web_ExtractStatusFromCurlResponse(web_CurlResponseLines() As String) As Long
    Dim web_StatusLineParts() As String

    web_StatusLineParts = VBA.Split(web_CurlResponseLines(web_FindStatusLine(web_CurlResponseLines)), " ")
    web_ExtractStatusFromCurlResponse = VBA.CLng(web_StatusLineParts(1))
End Function

Private Function web_ExtractStatusTextFromCurlResponse(web_CurlResponseLines() As String) As String
    Dim web_StatusLineParts() As String
    Dim web_i As Long
    Dim web_StatusText As String

    web_StatusLineParts = VBA.Split(web_CurlResponseLines(web_FindStatusLine(web_CurlResponseLines)), " ", 3)
    web_ExtractStatusTextFromCurlResponse = web_StatusLineParts(2)
End Function

Private Function web_ExtractResponseTextFromCurlResponse(web_CurlResponseLines() As String) As String
    Dim web_BlankLineIndex As Long
    Dim web_BodyLines() As String
    Dim web_WriteIndex As Long
    Dim web_ReadIndex As Long

    ' Find blank line before body
    web_BlankLineIndex = web_FindBlankLine(web_CurlResponseLines)

    ' Extract body string
    ReDim web_BodyLines(0 To UBound(web_CurlResponseLines) - web_BlankLineIndex - 1)

    web_WriteIndex = 0
    For web_ReadIndex = web_BlankLineIndex + 1 To UBound(web_CurlResponseLines)
        web_BodyLines(web_WriteIndex) = web_CurlResponseLines(web_ReadIndex)
        web_WriteIndex = web_WriteIndex + 1
    Next web_ReadIndex

    web_ExtractResponseTextFromCurlResponse = VBA.Join$(web_BodyLines, web_CrLf)
End Function

Private Function web_ExtractHeadersFromCurlResponse(web_CurlResponseLines() As String) As String
    Dim web_StatusLineIndex As Long
    Dim web_BlankLineIndex As Long
    Dim web_HeaderLines() As String
    Dim web_WriteIndex As Long
    Dim web_ReadIndex As Long

    ' Find status line and blank line before body
    web_StatusLineIndex = web_FindStatusLine(web_CurlResponseLines)
    web_BlankLineIndex = web_FindBlankLine(web_CurlResponseLines)

    ' Extract headers string
    ReDim web_HeaderLines(0 To web_BlankLineIndex - 2 - web_StatusLineIndex)

    web_WriteIndex = 0
    For web_ReadIndex = (web_StatusLineIndex + 1) To web_BlankLineIndex - 1
        web_HeaderLines(web_WriteIndex) = web_CurlResponseLines(web_ReadIndex)
        web_WriteIndex = web_WriteIndex + 1
    Next web_ReadIndex

    web_ExtractHeadersFromCurlResponse = VBA.Join$(web_HeaderLines, web_CrLf)
End Function

Private Function web_FindStatusLine(web_CurlResponseLines() As String) As Long
    ' Special case for cURL: 100 Continue is included before final status code
    ' -> ignore 100 and find final status code (next non-100 status code)
    For web_FindStatusLine = LBound(web_CurlResponseLines) To UBound(web_CurlResponseLines)
        If VBA.Trim$(web_CurlResponseLines(web_FindStatusLine)) <> "" Then
            If VBA.Split(web_CurlResponseLines(web_FindStatusLine), " ")(1) <> "100" Then
                Exit Function
            End If
        End If
    Next web_FindStatusLine
End Function

Private Function web_FindBlankLine(web_CurlResponseLines() As String) As Long
    For web_FindBlankLine = (web_FindStatusLine(web_CurlResponseLines) + 1) To UBound(web_CurlResponseLines)
        If VBA.Trim$(web_CurlResponseLines(web_FindBlankLine)) = "" Then
            Exit Function
        End If
    Next web_FindBlankLine
End Function
