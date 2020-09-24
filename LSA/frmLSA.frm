VERSION 5.00
Begin VB.Form frmLSA 
   Caption         =   "User Rights"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpenPolicy 
      Caption         =   "Open Policy"
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdClearRightsList 
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdQueryUserRights 
      Caption         =   "Query User Rights"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox lstRights 
      Height          =   1230
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   4815
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddToList 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.ListBox lstAccessRights 
      Height          =   1815
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ListBox lstUsers 
      Height          =   1035
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
   End
   Begin VB.ComboBox cboAccessRights 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtComputerName 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtAccountName 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Change List"
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Access Rights"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Users with Specified Access Right"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Account Name Access Rights"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblComputerName 
      Caption         =   "Computer Name"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblAccountName 
      Caption         =   "Account Name"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmLSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'COPYRIGHT JON Kelly 1999

Dim frm_PolicyHandle  As Long
Dim frm_lRetVal As Long
Dim frm_UnicodeBuffer    As LSA_UNICODE_STRING
Dim frm_ObjectAttributes As LSA_OBJECT_ATTRIBUTES
Dim frm_DesiredAccess As Long
Dim frm_lpMultiByteStr As String

Private Function GetLSAError(ByVal ErrorNumber As Long) As String

    Dim lReturn As Long
    
    lReturn = LsaNtStatusToWinError(ErrorNumber)

    If lReturn = ERROR_MR_MID_NOT_FOUND Then
        GetLSAError = ErrorNumber & ": " & lReturn & " - LSA ERROR NOT FOUND"
    Else
        GetLSAError = ErrorNumber & ": " & lReturn & " - " & MessageText(lReturn)
    End If

End Function

Public Function MessageText(ByVal lCode As Long) As String
    
    On Error Resume Next
    
    Dim sRtrnCode As String
    Dim lRet As Long
    
    sRtrnCode = Space$(256)
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, lCode, 0&, sRtrnCode, 256&, 0&)
    
    If lRet > 0 Then
        MessageText = Left(sRtrnCode, lRet)
    Else
        MessageText = "Error not found."
    End If
    
End Function

Private Sub cmdOpenPolicy_Click()
    cmdOpenPolicy.Enabled = False
    cmdAdd.Enabled = True
    cmdAddToList.Enabled = True
    cmdClearRightsList.Enabled = True
    cmdQueryUserRights.Enabled = True
    cmdRemove.Enabled = True
    OpenPolicy
End Sub

Private Sub Form_Load()

    txtAccountName.Text = Environ$("USERDOMAIN") & "\" & Environ$("USERNAME")
    txtComputerName.Text = Environ$("COMPUTERNAME")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim lRetVal As Long
    
    If frm_PolicyHandle <> 0 Then
        lRetVal = LsaClose(frm_PolicyHandle)
    
        'MODIFIED BY ZEILO
        If lRetVal <> STATUS_SUCCESS Then
            MsgBox GetLSAError(frm_lRetVal), vbCritical, "ERROR"
        End If
    
    End If
    
    frm_PolicyHandle = 0
    
End Sub

Private Sub cboAccessRights_Click()
    
    Dim lpMultiByteStr As String
    Dim userRights As LSA_UNICODE_STRING
    Dim lEnumerationBuffer As Long
    Dim lCountOfSIDs As Long
    Dim lRetVal As Long
    Dim lProcessHandle  As Long
    Dim memtest As pSidArray
    Dim lpNumberOfBytesWritten As Long
    Dim rdnName As ReferenceDomainName
    Dim rdnName2 As ReferenceDomainName
    Dim lReferencedDomain As Long
    Dim lReferencedDomain2 As Long
    Dim sReferencedDomainName As String
    Dim lUse As Long
    Dim nCountSids As Integer

    lstUsers.Clear
    lpMultiByteStr = cboAccessRights.Text
    CreateUnicodeString lpMultiByteStr, userRights
    lRetVal = LsaEnumerateAccountsWithUserRight(frm_PolicyHandle, _
                 userRights, lEnumerationBuffer, lCountOfSIDs)
    
    'MODIFIED BY ZEILO
    If lRetVal = STATUS_SUCCESS Then
        lProcessHandle = GetCurrentProcess()
        lRetVal = ReadProcessMemory(lProcessHandle, lEnumerationBuffer, _
                     memtest, 40, lpNumberOfBytesWritten)
    
        For nCountSids = 1 To lCountOfSIDs
            lReferencedDomain = 100
            lReferencedDomain2 = 100
            lRetVal = LookupAccountSid(vbNullString, _
                         memtest.sidData(nCountSids - 1), rdnName2, _
                         lReferencedDomain2, rdnName, lReferencedDomain, _
                         lUse)
            
            If lRetVal = 1 Then
                sReferencedDomainName = GetStringFromByteArray(rdnName.nameData) & "\" & GetStringFromByteArray(rdnName2.nameData)
            Else
                sReferencedDomainName = "Unknown Account"
            End If
            
            lstUsers.AddItem sReferencedDomainName
        Next nCountSids
    
        lRetVal = LsaFreeMemory(lEnumerationBuffer)
        
        'MODIFIED BY ZEILO
        If lRetVal <> STATUS_SUCCESS Then
            MsgBox GetLSAError(frm_lRetVal), vbCritical, "ERROR"
        End If
    
    ElseIf lRetVal = STATUS_NO_MORE_ENTRIES Then
        MsgBox "No accounts with this privilege.", vbInformation, "Info"
    Else
        MsgBox GetLSAError(frm_lRetVal), vbCritical, "ERROR"
    End If
    
End Sub

Public Function GetStringFromByteArray(bytArray() As Byte) As String
    Dim nChars As Integer
    For nChars = 0 To 257
        If bytArray(nChars) <> 0 And bytArray(nChars) <> 13 And _
               bytArray(nChars) <> 10 Then
            GetStringFromByteArray = GetStringFromByteArray & _
               Chr(bytArray(nChars))
        Else
            Exit For
        End If
    Next nChars
End Function

Public Function GetStringFromUnicodeByteArray(bytArray() As Byte) _
   As String
    Dim nChars As Integer

    For nChars = 0 To 257 Step 2
        If bytArray(nChars) <> 13 And bytArray(nChars) <> 10 Then
            GetStringFromUnicodeByteArray = _
               GetStringFromUnicodeByteArray & Chr(bytArray(nChars))
        Else
            Exit For
        End If
    Next nChars
End Function

Private Sub CreateUnicodeString(ByVal lpMultiByteStr As String, _
   UnicodeBuffer As LSA_UNICODE_STRING)
    Dim cchMultiByte As Long
    Dim cchWideChar As Long
    cchMultiByte = Len(lpMultiByteStr)
    UnicodeBuffer.Length = cchMultiByte * 2
    UnicodeBuffer.MaximumLength = UnicodeBuffer.Length + 2
    UnicodeBuffer.Buffer = String(UnicodeBuffer.MaximumLength, " ")
    cchWideChar = UnicodeBuffer.Length
    Dim lRetVal As Long
    lRetVal = MultiByteToWideChar(CP_ACP, 0, lpMultiByteStr, _
       cchMultiByte, UnicodeBuffer.Buffer, cchWideChar)
End Sub

Private Sub cmdAdd_Click()
    ChangeAccessRights True
End Sub

Private Sub cmdAddToList_Click()
    lstAccessRights.AddItem cboAccessRights.Text
End Sub

Private Sub cmdClearRightsList_Click()
    lstAccessRights.Clear
End Sub

Private Sub cmdQueryUserRights_Click()
    Dim sAccountName As String
    Dim lReferencedDomain As Long
    Dim rdnName As ReferenceDomainName
    Dim lSid As Long
    Dim pSidData As psid
    Dim nUse As Long
    Dim lRetVal As Long
    Dim lEnumerationBuffer As Long
    Dim lCountOfSIDs As Long
    Dim pData As lsaArray
    Dim lProcessHandle  As Long
    Dim lpNumberOfBytesWritten As Long
    lstRights.Clear
    sAccountName = txtAccountName.Text
    lReferencedDomain = 16
    lSid = 128
    lRetVal = LookupAccountName(vbNullString, sAccountName, pSidData, _
       lSid, rdnName, lReferencedDomain, nUse)

    If lRetVal <> 1 Then
        MsgBox "Invalid account.", vbExclamation, "Attention"
        Exit Sub
    End If

    lRetVal = LsaEnumerateAccountRights(frm_PolicyHandle, pSidData, _
       lEnumerationBuffer, lCountOfSIDs)

    'MODIFIED BY ZEILO
    If lRetVal = STATUS_SUCCESS Then '0 Then
        lProcessHandle = GetCurrentProcess()
        lRetVal = ReadProcessMemory2(lProcessHandle, lEnumerationBuffer, _
           pData, (60 * lCountOfSIDs) + 24, lpNumberOfBytesWritten)
        LoadStringsFromData pData, lCountOfSIDs
        lRetVal = LsaFreeMemory(lEnumerationBuffer)
        'MODIFIED BY ZEILO
        If lRetVal <> STATUS_SUCCESS Then
            MsgBox GetLSAError(frm_lRetVal), vbCritical, "ERROR"
        End If
    Else
        MsgBox GetLSAError(lRetVal), vbCritical, "ERROR"
    End If

End Sub

Private Sub cmdRemove_Click()
    ChangeAccessRights False
End Sub

Private Sub LoadStringsFromData(lpBuffer As lsaArray, ByVal _
   nNumberOfStrings)
    Dim nChars As Integer
    Dim nCount As Integer
    Dim nStringsCount As Integer
    Dim nStringsCount2 As Integer
    Dim sStrings(100) As String

    lstRights.Clear
    nStringsCount = 1

    For nChars = (nNumberOfStrings * 8) To 1000 Step 2
        If (Not (lpBuffer.lsaData(nChars + 2) = Asc("S") And _
           lpBuffer.lsaData(nChars + 4) = Asc("e") And _
           lpBuffer.lsaData(nChars + 6) < Asc("a"))) And _
           lpBuffer.lsaData(nChars) <> 0 And _
           lpBuffer.lsaData(nChars) <> 13 And _
           lpBuffer.lsaData(nChars) <> 10 Then '(Not _
           (lpBuffer.lsaData(nChars + 2) = Asc("S") And _
           lpBuffer.lsaData(nChars + 4) = Asc("e"))) And
           sStrings(nStringsCount2) = sStrings(nStringsCount2) & _
           Chr(lpBuffer.lsaData(nChars))
        Else

            If (lpBuffer.lsaData(nChars + 2) = Asc("S") And _
               lpBuffer.lsaData(nChars + 4) = Asc("e") And _
               lpBuffer.lsaData(nChars + 6) < Asc("a")) Then
                If lpBuffer.lsaData(nChars) <> 0 And _
                   lpBuffer.lsaData(nChars) <> 13 And _
                   lpBuffer.lsaData(nChars) <> 10 Then
                     sStrings(nStringsCount2) = _
                     sStrings(nStringsCount2) & _
                     Chr(lpBuffer.lsaData(nChars))
                End If
            End If

            If Len(sStrings(nStringsCount2)) <> 0 Then
                lstRights.AddItem sStrings(nStringsCount2)
            End If

            nStringsCount2 = nStringsCount2 + 1
            If nStringsCount2 >= nNumberOfStrings Then
                Exit For
            End If
        End If
    Next nChars
End Sub

Private Sub ChangeAccessRights(ByVal bAdd As Boolean)
    Dim sAccountName As String
    Dim lReferencedDomain As Long
    Dim lSid As Long
    Dim pSidData As psid
    Dim nUse As Long
    Dim rdnName As ReferenceDomainName
    Dim sReferencedDomainName As String
    Dim nCount As Integer
    Dim CountOfRights As Long
    Dim userRights As LSA_UNICODE_STRING

    sAccountName = txtAccountName.Text
    lReferencedDomain = 16
    lSid = 128
    frm_lRetVal = LookupAccountName(vbNullString, sAccountName, _
                  pSidData, lSid, rdnName, lReferencedDomain, nUse)
    
    If frm_lRetVal <> 1 Then
        MsgBox "Invalid account.", vbExclamation, "Attention"
        Exit Sub
    End If
    
    frm_lRetVal = 0
    sReferencedDomainName = GetStringFromByteArray(rdnName.nameData)
    CountOfRights = 1

    For nCount = 1 To lstAccessRights.ListCount
        If lstAccessRights.ListCount = 0 Then
            Exit For
        End If

        frm_lpMultiByteStr = lstAccessRights.List(nCount - 1)
        CreateUnicodeString frm_lpMultiByteStr, userRights

        If bAdd = True Then
            frm_lRetVal = LsaAddAccountRights(frm_PolicyHandle, _
                          pSidData, userRights, CountOfRights)
            'MODIFIED BY ZEILO
            If frm_lRetVal <> STATUS_SUCCESS Then
                MsgBox GetLSAError(frm_lRetVal), vbCritical, "ERROR"
            End If
            
        Else
            frm_lRetVal = LsaRemoveAccountRights(frm_PolicyHandle, _
                          pSidData, 0, userRights, CountOfRights)
            'MODIFIED BY ZEILO
            If frm_lRetVal <> STATUS_SUCCESS Then
                MsgBox GetLSAError(frm_lRetVal), vbCritical, "ERROR"
            End If
            
        End If
    Next nCount
    
End Sub

Private Sub OpenPolicy()
    
    frm_lpMultiByteStr = txtComputerName.Text
    CreateUnicodeString frm_lpMultiByteStr, frm_UnicodeBuffer
    'MODIFIED BY ZEILO
    frm_DesiredAccess = POLICY_ALL_ACCESS '2064
    'Open Policy
    frm_lRetVal = LsaOpenPolicy(frm_UnicodeBuffer, _
                  frm_ObjectAttributes, frm_DesiredAccess, _
                  frm_PolicyHandle)
                  
    If frm_lRetVal <> STATUS_SUCCESS Then '0 Then
        MsgBox GetLSAError(frm_lRetVal), vbCritical, "ERROR"
        cmdAdd.Enabled = False
        cmdAddToList.Enabled = False
        cmdClearRightsList.Enabled = False
        cmdQueryUserRights.Enabled = False
        cmdRemove.Enabled = False
        cmdOpenPolicy.Enabled = True
        Exit Sub
    End If

    'Fill privileges
    cboAccessRights.AddItem "SeInteractiveLogonRight"
    cboAccessRights.AddItem "SeNetworkLogonRight"
    cboAccessRights.AddItem "SeBatchLogonRight"
    cboAccessRights.AddItem "SeServiceLogonRight"
    cboAccessRights.AddItem "SeDenyInteractiveLogonRight"
    cboAccessRights.AddItem "SeDenyNetworkLogonRight"
    cboAccessRights.AddItem "SeDenyBatchLogonRight"
    cboAccessRights.AddItem "SeDenyServiceLogonRight"
    cboAccessRights.AddItem "SeCreateTokenPrivilege"
    cboAccessRights.AddItem "SeAssignPrimaryTokenPrivilege"
    cboAccessRights.AddItem "SeLockMemoryPrivilege"
    cboAccessRights.AddItem "SeIncreaseQuotaPrivilege"
    cboAccessRights.AddItem "SeUnsolicitedInputPrivilege"
    cboAccessRights.AddItem "SeMachineAccountPrivilege"
    cboAccessRights.AddItem "SeTcbPrivilege"
    cboAccessRights.AddItem "SeSecurityPrivilege"
    cboAccessRights.AddItem "SeTakeOwnershipPrivilege"
    cboAccessRights.AddItem "SeLoadDriverPrivilege"
    cboAccessRights.AddItem "SeSystemProfilePrivilege"
    cboAccessRights.AddItem "SeSystemtimePrivilege"
    cboAccessRights.AddItem "SeProfileSingleProcessPrivilege"
    cboAccessRights.AddItem "SeIncreaseBasePriorityPrivilege"
    cboAccessRights.AddItem "SeCreatePagefilePrivilege"
    cboAccessRights.AddItem "SeCreatePermanentPrivilege"
    cboAccessRights.AddItem "SeBackupPrivilege"
    cboAccessRights.AddItem "SeRestorePrivilege"
    cboAccessRights.AddItem "SeShutdownPrivilege"
    cboAccessRights.AddItem "SeDebugPrivilege"
    cboAccessRights.AddItem "SeAuditPrivilege"
    cboAccessRights.AddItem "SeSystemEnvironmentPrivilege"
    cboAccessRights.AddItem "SeChangeNotifyPrivilege"
    cboAccessRights.AddItem "SeRemoteShutdownPrivilege"
    'Select desired privilege
    cboAccessRights.Text = "SeServiceLogonRight"
    
End Sub
