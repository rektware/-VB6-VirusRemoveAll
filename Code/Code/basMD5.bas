Attribute VB_Name = "basMD5"
'This module is used to gather the contents of a file quickly and to grab the MD5 of a file quickly by using API functions. Use this
'code in any projects you wish, no need to give credit. Please vote though.

'marcin@malwarebytes.org if you have any questions.

Option Explicit

Public Const OPEN_EXISTING = 3
Public Const GENERIC_READ = &H80000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const BUFFER_SIZE As Long = 255

'MD5 Hashing
Public Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"
Public Const MS_BASE_PROV As String = "Microsoft Base Cryptographic Provider v1.0"
Public Const PROV_RSA_FULL As Long = 1
Public Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Public Const ALG_TYPE_STREAM As Long = 2048
Public Const ALG_TYPE_ANY As Long = 0
Public Const ALG_SID_RC4 As Long = 1
Public Const ALG_SID_MD5 As Long = 3
Public Const CALG_RC4 As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4
Public Const CRYPT_VERIFYCONTEXT = &HF0000000
Public Const CRYPT_NEWKEYSET As Long = 8
Public Const ENCRYPT_ALGORITHM As Long = CALG_RC4
Public Const ENCRYPT_NUMBERKEY As String = "16006833"
Public Const ALG_CLASS_HASH As Long = 32768
Public Const HP_HASHVAL As Long = 2
Public Const HP_HASHSIZE As Long = 4

'Faster hashing
Public Const HASH_TYPE = ALG_TYPE_ANY Or ALG_CLASS_HASH Or ALG_SID_MD5

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

'MD5 Hashing
Public Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Public Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Public Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Public Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Public Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Public Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Public Function InputFile$(ByRef sFile$)
    Dim hFile&, uBuffer() As Byte, lFileSize&, lBytesRead&
    
    'Get a handle to the file
    hFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_EXISTING, 0, 0)
        
    'Can't open file
    If hFile = -1 Then Exit Function
    lFileSize = GetFileSize(hFile, 0)
    
    If lFileSize < 1 Then
        CloseHandle hFile
            Exit Function
    End If
    
    'Prepare the buffer
    ReDim uBuffer(lFileSize - 1)
    
    'Read the file
    If ReadFile(hFile, uBuffer(0), lFileSize, lBytesRead, ByVal 0) <> 0 Then
        If lBytesRead <> lFileSize Then
            ReDim Preserve uBuffer(lBytesRead)
        End If
    
        InputFile = StrConv(uBuffer, vbUnicode)
    End If
        
    'Close the handle to the file
    CloseHandle hFile
End Function

Public Function GetMD5$(ByRef sFileName$)
    'Get the MD5
    GetMD5 = MD5String(InputFile(sFileName))
End Function

Public Function MD5String(ByRef sString$)
    Dim uMD5() As Byte, lMD5Len&, i&, sMD5$, hCrypt&, hHash&
    
        'Prepare the byte array
        ReDim uMD5(BUFFER_SIZE)
    
        'Acquire the MD5 hash generator
        If CryptAcquireContext(hCrypt, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
            If CryptCreateHash(hCrypt, HASH_TYPE, 0, 0, hHash) <> 0 Then
                If CryptHashData(hHash, sString, Len(sString), 0) <> 0 Then
                    If CryptGetHashParam(hHash, HP_HASHSIZE, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then
                        lMD5Len = uMD5(0)
                    
                        If CryptGetHashParam(hHash, HP_HASHVAL, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then
                            For i = 0 To lMD5Len - 1
                                sMD5 = sMD5 & (Right$("0" & Hex$(uMD5(i)), 2))
                            Next i
                            
                            MD5String = sMD5
                        End If
                    End If
                End If
            End If
        End If
        
    'Destroy the MD5 hash generator
    CryptDestroyHash hHash
    CryptReleaseContext hCrypt, 0
End Function


