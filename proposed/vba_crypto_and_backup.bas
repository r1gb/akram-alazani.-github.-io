' Module: modCryptoAndBackup
Option Compare Database
Option Explicit

' --- SHA256 using CryptoAPI (Windows)
#If VBA7 Then
  Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
    (ByRef phProv As LongPtr, ByVal pszContainer As String, ByVal pszProvider As String, _
     ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

  Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long

  Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" _
    (ByVal hHash As LongPtr, ByRef pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

  Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" _
    (ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long

  Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As LongPtr) As Long
  Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
#Else
  Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
    (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
     ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

  Private Declare Function CryptCreateHash Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long

  Private Declare Function CryptHashData Lib "advapi32.dll" _
    (ByVal hHash As Long, ByRef pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

  Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
    (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long

  Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
  Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
#End If

Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA_256 As Long = &H800C
Private Const HP_HASHVAL As Long = &H2

Public Function SHA256Hex(ByVal txt As String) As String
    Dim hProv As LongPtr, hHash As LongPtr
    Dim ret As Long
    Dim b() As Byte
    Dim hashLen As Long
    Dim hashBytes() As Byte
    Dim i As Long
    Dim sHex As String

    ret = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)
    If ret = 0 Then Err.Raise vbObjectError + 1, , "CryptAcquireContext failed"

    ret = CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash)
    If ret = 0 Then
        CryptReleaseContext hProv, 0
        Err.Raise vbObjectError + 2, , "CryptCreateHash failed"
    End If

    ' Convert text to ANSI bytes (consider UTF-8 if needed)
    b = StrConv(txt, vbFromUnicode)
    ret = CryptHashData(hHash, b(0), UBound(b) + 1, 0)
    If ret = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Err.Raise vbObjectError + 3, , "CryptHashData failed"
    End If

    hashLen = 0
    ret = CryptGetHashParam(hHash, HP_HASHVAL, ByVal 0&, hashLen, 0)
    If ret = 0 And hashLen = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Err.Raise vbObjectError + 4, , "CryptGetHashParam failed to get length"
    End If

    ReDim hashBytes(hashLen - 1)
    ret = CryptGetHashParam(hHash, HP_HASHVAL, hashBytes(0), hashLen, 0)
    If ret = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Err.Raise vbObjectError + 5, , "CryptGetHashParam failed to get data"
    End If

    sHex = ""
    For i = 0 To UBound(hashBytes)
        sHex = sHex & Right("0" & Hex$(hashBytes(i)), 2)
    Next i

    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0

    SHA256Hex = LCase$(sHex)
End Function

' --- User authentication example (uses tblUsers)
Public Function AuthenticateUser(ByVal sUsername As String, ByVal sPassword As String) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim storedHash As String, storedSalt As String, calcHash As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT UserID, PasswordHash, PasswordSalt, IsActive FROM tblUsers WHERE Username = '" & Replace(sUsername, "'", "''") & "'")
    If rs.EOF Then
        AuthenticateUser = 0
    Else
        If Not rs!IsActive Then
            AuthenticateUser = 0
        Else
            storedHash = Nz(rs!PasswordHash, "")
            storedSalt = Nz(rs!PasswordSalt, "")
            calcHash = SHA256Hex(storedSalt & sPassword)
            If LCase(calcHash) = LCase(storedHash) Then
                AuthenticateUser = rs!UserID
                ' update last login and insert to tblLoginLog
                db.Execute "INSERT INTO tblLoginLog (UserID, LoginTime, Success, ComputerName) VALUES (" & rs!UserID & ",#" & Now() & "#, True, '" & Replace(Environ("COMPUTERNAME"), "'", "''") & "')"
                db.Execute "UPDATE tblUsers SET LastLogin = #" & Now() & "# WHERE UserID = " & rs!UserID
            Else
                AuthenticateUser = 0
                db.Execute "INSERT INTO tblLoginLog (UserID, LoginTime, Success, ComputerName, Notes) VALUES (" & rs!UserID & ",#" & Now() & "#, False, '" & Replace(Environ("COMPUTERNAME"), "'", "''") & "', 'Bad password')"
            End If
        End If
    End If
ExitHandler:
    On Error Resume Next
    rs.Close: Set rs = Nothing: Set db = Nothing
    Exit Function
ErrHandler:
    AuthenticateUser = 0
    Resume ExitHandler
End Function

' --- AutoBackup function: copies current .accdb to destination folder
Public Function AutoBackupDatabase(destFolder As String, Optional createdByUserID As Long = 0) As Boolean
    On Error GoTo ErrHandler
    Dim srcPath As String, destPath As String
    Dim fso As Object
    srcPath = CurrentDb.Name ' full path to current accdb
    If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(destFolder) Then fso.CreateFolder (destFolder)
    destPath = destFolder & "Backup_" & Format(Now(), "yyyyMMdd_HHmmss") & "_" & fso.GetFileName(srcPath)
    FileCopy srcPath, destPath
    ' سجل النسخة في tblBackups
    CurrentDb.Execute "INSERT INTO tblBackups (BackupPath, CreatedAt, CreatedBy) VALUES ('" & Replace(destPath, "'", "''") & "',#" & Now() & "#," & createdByUserID & ")"
    AutoBackupDatabase = True
    Exit Function
ErrHandler:
    AutoBackupDatabase = False
End Function
