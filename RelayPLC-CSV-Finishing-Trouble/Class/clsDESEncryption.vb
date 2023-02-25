
Imports System.Security.Cryptography

Public Class clsDESEncryption
    Private TripleDes As New TripleDESCryptoServiceProvider

    Sub New(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()
        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key. 
        Dim keyBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash. 
        ReDim Preserve hash(length - 1)
        Return hash
    End Function

    Public Function EncryptData(ByVal plaintext As String) As String
        ' Convert the plaintext string to a byte array. 
        Dim plaintextBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream. 
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream. 
        Dim encStream As New CryptoStream(ms,
            TripleDes.CreateEncryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string. 
        Return Convert.ToBase64String(ms.ToArray)
    End Function

    Public Function DecryptData(ByVal encryptedtext As String) As String
        ' Convert the encrypted text string to a byte array. 
        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

        ' Create the stream. 
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream. 
        Dim decStream As New CryptoStream(ms,
            TripleDes.CreateDecryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
        decStream.FlushFinalBlock()

        ' Convert the plaintext stream to a string. 
        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function

    Public Function fc_Encrypt(ByVal psFile As String) As String
        On Error GoTo err_handler
        '.MousePointer = vbHourglass

        Dim strChar As String
        Dim intCount As Integer
        Dim i As Integer = 1
        fc_Encrypt = ""


        Do
            If i > Len(psFile) Then
                strChar = ""
                Exit Do
            End If
            strChar = Mid(psFile, i, 1)
            fc_Encrypt = fc_Encrypt & Format(Asc(Mid(psFile, i, 1)) + 100, "000")
            i = i + 1
        Loop Until strChar = ""

err_exit:
        'Screen.MousePointer = vbDefault
        Return fc_Encrypt
        Exit Function
err_handler:
        'Screen.MousePointer = vbDefault
        'MsgBox("Err. Number : " & Err.number & vbNewLine & "Err. Description : " & Err.Description, vbCritical, "Error encrypting file")
        Err.Clear()
        Resume err_exit
    End Function

    Public Function fc_Decrypt(ByVal psFile As String) As String
        On Error GoTo err_handler
        'Screen.MousePointer = vbHourglass

        Dim strCode As String
        Dim strPos As Integer
        Dim strChar As String
        Dim i As Integer = 1
        fc_Decrypt = ""

        Do

            If i >= Len(psFile.Trim) Then
                strChar = ""
                Exit Do
            End If
            strChar = Mid(psFile, i, 3)
            fc_Decrypt = fc_Decrypt & Chr(CInt((Mid(psFile, i, 3)) - 100))
            i = i + 3
        Loop Until strChar = ""
        'Loop Until psFile = ""

err_exit:
        'Screen.MousePointer = vbDefault
        Exit Function
err_handler:
        'Screen.MousePointer = vbDefault
        'MsgBox("Err. Number : " & Err.Number & vbNewLine & "Err. Description : " & Err.Description, vbCritical, "Error decrypting file")
        Err.Clear()
        Resume err_exit
    End Function
End Class
