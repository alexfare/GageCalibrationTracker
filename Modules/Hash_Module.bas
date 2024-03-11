Attribute VB_Name = "Hash_Module"
Option Explicit

Public Function SHA512(sIn As String, Optional bB64 As Boolean = 0) As String

    Dim oT As Object, oSHA512 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
    
    TextToHash = oT.Getbytes_4(sIn)
    bytes = oSHA512.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA512 = ConvToBase64String(bytes)
    Else
       SHA512 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA512 = Nothing
    
End Function

Function StrToSHA512Salt(ByVal sIn As String, ByVal sSecretKey As String, _
                           Optional ByVal b64 As Boolean = False) As String
    'Returns a sha512 STRING HASH in function name, modified by the parameter sSecretKey.
    'This hash differs from that of SHA512 using the SHA512Managed class.
    'HMAC class inputs are hashed twice;first input and key are mixed before hashing,
    'then the key is mixed with the result and hashed again.
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SecretKey() As Byte
    Dim bytes() As Byte
    
    'Test results with both strings empty:
    '128 Hex:    b936cee86c9f...etc
    '88 Base-64:   uTbO6Gyfh6pd...etc
    
    'create text and crypto objects
    Set asc = CreateObject("System.Text.UTF8Encoding")
    
    'Any of HMACSHAMD5,HMACSHA1,HMACSHA256,HMACSHA384,or HMACSHA512 can be used
    'for corresponding hashes, albeit not matching those of Managed classes.
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA512")

    'make a byte array of the text to hash
    bytes = asc.Getbytes_4(sIn)
    'make a byte array of the private key
    SecretKey = asc.Getbytes_4(sSecretKey)
    'add the private key property to the encryption object
    enc.Key = SecretKey

    'make a byte array of the hash
    bytes = enc.ComputeHash_2((bytes))
    
    'convert the byte array to string
    If b64 = True Then
       StrToSHA512Salt = ConvToBase64String(bytes)
    Else
       StrToSHA512Salt = ConvToHexString(bytes)
    End If
    
    'release object variables
    Set asc = Nothing
    Set enc = Nothing

End Function

Private Function ConvToBase64String(vIn As Variant) As Variant
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
   
   Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Private Function ConvToHexString(vIn As Variant) As Variant
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

