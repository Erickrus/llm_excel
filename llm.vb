' ==========================================
' 1. CUSTOM FUNCTIONS (Accepts Both Text and Ranges)
' ==========================================

Function LLM(inputData As Variant) As String
    LLM = GenerateTriggerString("LLM", inputData)
End Function

Function CRAWL(inputData As Variant) As String
    CRAWL = GenerateTriggerString("CRAWL", inputData)
End Function

Function READ(inputData As Variant) As String
    READ = GenerateTriggerString("READ", inputData)
End Function

Function AGENT(inputData As Variant) As String
    AGENT = GenerateTriggerString("AGENT", inputData)
End Function

' Helper to create the unique trigger string safely
Function GenerateTriggerString(funcName As String, inputData As Variant) As String
    Dim refInfo As String
    
    If IsObject(inputData) Then
        ' It is a cell reference (e.g. A1)
        If Not inputData Is Nothing Then
            refInfo = inputData.Address(False, False)
        Else
            refInfo = "REF_ERR"
        End If
    Else
        ' It is a direct string/number (e.g. "Hello")
        refInfo = "VAL"
    End If
    
    GenerateTriggerString = "[" & funcName & " trigger for " & refInfo & "]"
End Function

' ==========================================
' 2. MAIN PROCESSING LOGIC
' ==========================================

Sub TriggerProcessing()
    Dim cell As Range
    Dim inputVal As String, formula As String
    Dim encodedInput As String, encodedResult As String, decodedResult As String
    Dim uuid As String, folder As String
    Dim reqPath As String, resPath As String
    Dim f As Integer
    Dim startTime As Double, timeoutSeconds As Double
    Dim ws As Worksheet
    Dim requestType As String
    
    ' Variables for parsing
    Dim formulas As Variant
    Dim argString As String
    Dim evalResult As Variant
    Dim p1 As Long, p2 As Long
    
    formulas = Array("LLM", "CRAWL", "READ", "AGENT")
        
    Set ws = ActiveSheet
    
    ' Use System Temp folder to avoid OneDrive/SharePoint path errors
    ' folder = Environ("TEMP") & Application.PathSeparator & "llm_temp"
    folder = ThisWorkbook.Path & Application.PathSeparator & "llm_temp"
    
    ' Create folder safely
    On Error Resume Next
    MkDir folder
    On Error GoTo 0
    
    timeoutSeconds = 80

    For Each cell In Selection
        formula = cell.formula
        requestType = ""
        
        Dim i As Long
        For i = LBound(formulas) To UBound(formulas)
            ' Check if formula starts with =FUNCTION(
            If UCase(Left(formula, Len(formulas(i)) + 2)) = "=" & formulas(i) & "(" Then
                requestType = formulas(i)
                
                ' Robust Parsing: Extract content between first ( and last )
                p1 = InStr(formula, "(")
                p2 = InStrRev(formula, ")")
                
                If p1 > 0 And p2 > p1 Then
                    argString = Mid(formula, p1 + 1, p2 - p1 - 1)
                    
                    ' Evaluate the argument to get the actual value or range object
                    On Error Resume Next
                    Set evalResult = Evaluate(argString) ' Try to get Object (Range)
                    If evalResult Is Nothing Then
                        evalResult = Evaluate(argString) ' Get Value (String/Number)
                    End If
                    On Error GoTo 0
                End If
                
                Exit For
            End If
        Next i
        
        If requestType <> "" Then
            inputVal = ""
            
            ' --- Determine Input Value (Range vs String) ---
            If IsObject(evalResult) Then
                ' It is a Range (e.g. B1)
                If Not evalResult Is Nothing Then
                    inputVal = evalResult.Value
                End If
            Else
                ' It is a literal (e.g. "hello")
                inputVal = CStr(evalResult)
            End If
            ' -----------------------------------------------

            If inputVal = "" Then
                cell.Offset(0, 1).Value = "[Error: empty input]"
                GoTo NextCell
            End If
            
            ' Generate unique ID and file paths
            uuid = Format(Now, "yyyymmddhhmmss") & Int(Rnd() * 10000)
            reqPath = folder & Application.PathSeparator & "request_" & uuid & ".txt"
            resPath = folder & Application.PathSeparator & "response_" & uuid & ".txt"
            
            ' 1. ENCODE input to Base64 (handles Chinese/UTF-8/Emojis)
            encodedInput = EncodeBase64(inputVal)
            
            ' Write request: TYPE,BASE64_STRING
            f = FreeFile
            Open reqPath For Output As #f
            Print #f, requestType & "," & encodedInput
            Close #f
            
            ' Wait for Python
            startTime = Timer
            Do While Dir(resPath) = ""
                DoEvents
                If Timer - startTime > timeoutSeconds Then
                    cell.Offset(0, 1).Value = "[Timeout waiting for response]"
                    GoTo NextCell
                End If
                ' Small wait to save CPU
                Application.Wait [=Now() + TimeValue("00:00:00.1")]
            Loop
            
            ' Read the response (It will be a Base64 string)
            f = FreeFile
            Open resPath For Input As #f
            Line Input #f, encodedResult
            Close #f
            
            ' 2. DECODE result from Base64 back to String
            decodedResult = DecodeBase64(encodedResult)
            
            ' Write result in adjacent cell
            cell.Offset(0, 1).Value = decodedResult
            
            ' Clean up
            If Dir(resPath) <> "" Then Kill resPath
        ElseIf requestType <> "" Then
            cell.Offset(0, 1).Value = "[Error: invalid reference]"
        End If
NextCell:
    Next cell
End Sub

' ==========================================
' 3. PURE VBA BASE64 ENCODER (HANDLES EMOJIS & UTF-8)
' ==========================================
Function EncodeBase64(text As String) As String
    Dim arrData() As Byte
    Dim arrResult() As Byte
    Dim strRef As String
    Dim i As Long, j As Long
    Dim c As Long, c2 As Long
    Dim scalar As Long
    Dim bCounter As Long
    Dim utf8Bytes() As Byte
    
    If Len(text) = 0 Then Exit Function
    
    ' 1. Convert String to UTF-8 Byte Array (Handling Surrogates for Emojis)
    ReDim utf8Bytes(Len(text) * 4) ' Max size (4 bytes per char)
    bCounter = 0
    
    i = 1
    Do While i <= Len(text)
        c = AscW(Mid(text, i, 1)) And &HFFFF& ' Unsigned
        
        ' Check for High Surrogate (Start of Emoji pair: D800 - DBFF)
        scalar = c
        If (c >= &HD800&) And (c <= &HDBFF&) And (i < Len(text)) Then
            c2 = AscW(Mid(text, i + 1, 1)) And &HFFFF&
            ' Check Low Surrogate (DC00 - DFFF)
            If (c2 >= &HDC00&) And (c2 <= &HDFFF&) Then
                ' Combine to Unicode Scalar
                scalar = &H10000 + ((c And &H3FF&) * &H400&) + (c2 And &H3FF&)
                i = i + 1 ' Skip next char
            End If
        End If
        
        ' Encode Scalar to UTF-8 Bytes
        If scalar < 128 Then
            utf8Bytes(bCounter) = scalar
            bCounter = bCounter + 1
        ElseIf scalar < 2048 Then
            utf8Bytes(bCounter) = &HC0 Or (scalar \ &H40&)
            bCounter = bCounter + 1
            utf8Bytes(bCounter) = &H80 Or (scalar And &H3F&)
            bCounter = bCounter + 1
        ElseIf scalar < 65536 Then
            utf8Bytes(bCounter) = &HE0 Or (scalar \ &H1000&)
            bCounter = bCounter + 1
            utf8Bytes(bCounter) = &H80 Or ((scalar \ &H40&) And &H3F&)
            bCounter = bCounter + 1
            utf8Bytes(bCounter) = &H80 Or (scalar And &H3F&)
            bCounter = bCounter + 1
        Else ' 4-Byte (Emojis)
            utf8Bytes(bCounter) = &HF0 Or (scalar \ &H40000)
            bCounter = bCounter + 1
            utf8Bytes(bCounter) = &H80 Or ((scalar \ &H1000&) And &H3F&)
            bCounter = bCounter + 1
            utf8Bytes(bCounter) = &H80 Or ((scalar \ &H40&) And &H3F&)
            bCounter = bCounter + 1
            utf8Bytes(bCounter) = &H80 Or (scalar And &H3F&)
            bCounter = bCounter + 1
        End If
        
        i = i + 1
    Loop
    
    If bCounter = 0 Then Exit Function
    ReDim Preserve utf8Bytes(bCounter - 1)
    
    ' 2. Convert UTF-8 Bytes to Base64 String
    strRef = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim outLen As Long
    outLen = ((UBound(utf8Bytes) + 1 + 2) \ 3) * 4
    ReDim arrResult(outLen - 1)
    
    Dim b1 As Integer, b2 As Integer, b3 As Integer
    Dim enc1 As Integer, enc2 As Integer, enc3 As Integer, enc4 As Integer
    
    j = 0
    For i = 0 To UBound(utf8Bytes) Step 3
        b1 = utf8Bytes(i)
        If i + 1 <= UBound(utf8Bytes) Then b2 = utf8Bytes(i + 1) Else b2 = 0
        If i + 2 <= UBound(utf8Bytes) Then b3 = utf8Bytes(i + 2) Else b3 = 0
        
        enc1 = b1 \ 4
        enc2 = ((b1 And 3) * 16) Or (b2 \ 16)
        enc3 = ((b2 And 15) * 4) Or (b3 \ 64)
        enc4 = b3 And 63
        
        arrResult(j) = Asc(Mid(strRef, enc1 + 1, 1)): j = j + 1
        arrResult(j) = Asc(Mid(strRef, enc2 + 1, 1)): j = j + 1
        
        If i + 1 > UBound(utf8Bytes) Then
            arrResult(j) = 61 ' =
        Else
            arrResult(j) = Asc(Mid(strRef, enc3 + 1, 1))
        End If
        j = j + 1
        
        If i + 2 > UBound(utf8Bytes) Then
            arrResult(j) = 61 ' =
        Else
            arrResult(j) = Asc(Mid(strRef, enc4 + 1, 1))
        End If
        j = j + 1
    Next i
    
    EncodeBase64 = StrConv(arrResult, vbUnicode)
End Function

' ==========================================
' 4. PURE VBA BASE64 DECODER (HANDLES EMOJIS & UTF-8)
' ==========================================
Function DecodeBase64(base64Text As String) As String
    Dim b64Map(255) As Integer
    Dim b64Str As String
    Dim i As Long, j As Long
    Dim b1 As Integer, b2 As Integer, b3 As Integer, b4 As Integer
    Dim outBytes() As Byte
    Dim cleanB64 As String
    
    If Len(base64Text) = 0 Then Exit Function
    
    ' 1. Initialize Map
    b64Str = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    For i = 0 To 255: b64Map(i) = -1: Next i
    For i = 1 To 64
        b64Map(Asc(Mid(b64Str, i, 1))) = i - 1
    Next i
    
    ' 2. Clean Input
    cleanB64 = Replace(Replace(Replace(base64Text, vbCr, ""), vbLf, ""), " ", "")
    Dim inputLen As Long
    inputLen = Len(cleanB64)
    
    ' 3. Decode Base64 to Byte Array
    ReDim outBytes((inputLen \ 4) * 3)
    j = 0
    For i = 1 To inputLen Step 4
        If i + 3 > inputLen Then Exit For ' Safety check
        
        b1 = b64Map(Asc(Mid(cleanB64, i, 1)))
        b2 = b64Map(Asc(Mid(cleanB64, i + 1, 1)))
        
        If Mid(cleanB64, i + 2, 1) = "=" Then b3 = -1 Else b3 = b64Map(Asc(Mid(cleanB64, i + 2, 1)))
        If Mid(cleanB64, i + 3, 1) = "=" Then b4 = -1 Else b4 = b64Map(Asc(Mid(cleanB64, i + 3, 1)))
        
        If b1 = -1 Or b2 = -1 Then Exit For
        
        outBytes(j) = (b1 * 4) Or (b2 \ 16)
        j = j + 1
        
        If b3 <> -1 Then
            outBytes(j) = ((b2 And 15) * 16) Or (b3 \ 4)
            j = j + 1
        End If
        
        If b4 <> -1 Then
            outBytes(j) = ((b3 And 3) * 64) Or b4
            j = j + 1
        End If
    Next i
    
    If j > 0 Then ReDim Preserve outBytes(j - 1) Else Exit Function
    
    ' 4. Convert UTF-8 Bytes to VBA String (Handling Emojis)
    Dim finalStr As String
    Dim charCode As Long
    Dim idx As Long
    Dim byte1 As Integer
    
    idx = 0
    Do While idx <= UBound(outBytes)
        byte1 = outBytes(idx)
        
        If byte1 < 128 Then
            ' 1 Byte (ASCII)
            finalStr = finalStr & ChrW(byte1)
            idx = idx + 1
            
        ElseIf (byte1 And &HE0) = &HC0 Then
            ' 2 Bytes
            If idx + 1 > UBound(outBytes) Then Exit Do
            charCode = (CLng(byte1 And &H1F) * &H40&) Or (outBytes(idx + 1) And &H3F&)
            finalStr = finalStr & ChrW(charCode)
            idx = idx + 2
            
        ElseIf (byte1 And &HF0) = &HE0 Then
            ' 3 Bytes (Typical Chinese/Asian)
            If idx + 2 > UBound(outBytes) Then Exit Do
            charCode = (CLng(byte1 And &HF) * &H1000&) Or _
                       (CLng(outBytes(idx + 1) And &H3F) * &H40&) Or _
                       (outBytes(idx + 2) And &H3F)
            finalStr = finalStr & ChrW(charCode)
            idx = idx + 3
            
        ElseIf (byte1 And &HF8) = &HF0 Then
            ' 4 Bytes (Emojis / Symbols)
            If idx + 3 > UBound(outBytes) Then Exit Do
            
            charCode = (CLng(byte1 And &H7) * &H40000) Or _
                       (CLng(outBytes(idx + 1) And &H3F) * &H1000&) Or _
                       (CLng(outBytes(idx + 2) And &H3F) * &H40&) Or _
                       (outBytes(idx + 3) And &H3F)
            
            ' Surrogate Pair
            charCode = charCode - &H10000
            Dim highSurrogate As Long
            Dim lowSurrogate As Long
            
            highSurrogate = &HD800& + (charCode \ &H400&)
            lowSurrogate = &HDC00& + (charCode And &H3FF&)
            
            finalStr = finalStr & ChrW(highSurrogate) & ChrW(lowSurrogate)
            idx = idx + 4
        Else
            idx = idx + 1
        End If
    Loop
    
    DecodeBase64 = finalStr
End Function

