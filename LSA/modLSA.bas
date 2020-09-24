Attribute VB_Name = "modLSA"
Option Explicit

'COPYRIGHT JON Kelly 1999

'INCLUDED BY ZEILO
'---------------------------------------------------------------------------------------------
Public Const POLICY_AUDIT_EVENT_FAILURE As Long = &H2
Public Const POLICY_AUDIT_EVENT_NONE As Long = &H4
Public Const POLICY_AUDIT_EVENT_SUCCESS As Long = &H1
Public Const POLICY_AUDIT_EVENT_UNCHANGED As Long = &H0
Public Const POLICY_AUDIT_LOG_ADMIN As Long = &H200&
Public Const POLICY_CREATE_ACCOUNT As Long = &H10&
Public Const POLICY_CREATE_PRIVILEGE As Long = &H40&
Public Const POLICY_CREATE_SECRET As Long = &H20&
Public Const POLICY_ERRV_CRAZY_FLOWSPEC As Long = 57
Public Const POLICY_ERRV_EXPIRED_CREDENTIALS As Long = 4
Public Const POLICY_ERRV_EXPIRED_USER_TOKEN As Long = 51
Public Const POLICY_ERRV_GLOBAL_DEF_FLOW_COUNT As Long = 1
Public Const POLICY_ERRV_GLOBAL_DEF_FLOW_DURATION As Long = 9
Public Const POLICY_ERRV_GLOBAL_DEF_FLOW_RATE As Long = 17
Public Const POLICY_ERRV_GLOBAL_DEF_PEAK_RATE As Long = 25
Public Const POLICY_ERRV_GLOBAL_DEF_SUM_FLOW_RATE As Long = 33
Public Const POLICY_ERRV_GLOBAL_DEF_SUM_PEAK_RATE As Long = 41
Public Const POLICY_ERRV_GLOBAL_GRP_FLOW_COUNT As Long = 2
Public Const POLICY_ERRV_GLOBAL_GRP_FLOW_DURATION As Long = 10
Public Const POLICY_ERRV_GLOBAL_GRP_FLOW_RATE As Long = 18
Public Const POLICY_ERRV_GLOBAL_GRP_PEAK_RATE As Long = 26
Public Const POLICY_ERRV_GLOBAL_GRP_SUM_FLOW_RATE As Long = 34
Public Const POLICY_ERRV_GLOBAL_GRP_SUM_PEAK_RATE As Long = 42
Public Const POLICY_ERRV_GLOBAL_UNAUTH_USER_FLOW_COUNT As Long = 4
Public Const POLICY_ERRV_GLOBAL_UNAUTH_USER_FLOW_DURATION As Long = 12
Public Const POLICY_ERRV_GLOBAL_UNAUTH_USER_FLOW_RATE As Long = 20
Public Const POLICY_ERRV_GLOBAL_UNAUTH_USER_PEAK_RATE As Long = 28
Public Const POLICY_ERRV_GLOBAL_UNAUTH_USER_SUM_FLOW_RATE As Long = 36
Public Const POLICY_ERRV_GLOBAL_UNAUTH_USER_SUM_PEAK_RATE As Long = 44
Public Const POLICY_ERRV_GLOBAL_USER_FLOW_COUNT As Long = 3
Public Const POLICY_ERRV_GLOBAL_USER_FLOW_DURATION As Long = 11
Public Const POLICY_ERRV_GLOBAL_USER_FLOW_RATE As Long = 19
Public Const POLICY_ERRV_GLOBAL_USER_PEAK_RATE As Long = 27
Public Const POLICY_ERRV_GLOBAL_USER_SUM_FLOW_RATE As Long = 35
Public Const POLICY_ERRV_GLOBAL_USER_SUM_PEAK_RATE As Long = 43
Public Const POLICY_ERRV_IDENTITY_CHANGED As Long = 5
Public Const POLICY_ERRV_INSUFFICIENT_PRIVILEGES As Long = 3
Public Const POLICY_ERRV_NO_ACCEPTS As Long = 55
Public Const POLICY_ERRV_NO_MEMORY As Long = 56
Public Const POLICY_ERRV_NO_MORE_INFO As Long = 1
Public Const POLICY_ERRV_NO_PRIVILEGES As Long = 50
Public Const POLICY_ERRV_NO_RESOURCES As Long = 52
Public Const POLICY_ERRV_PRE_EMPTED As Long = 53
Public Const POLICY_ERRV_SUBNET_DEF_FLOW_COUNT As Long = 5
Public Const POLICY_ERRV_SUBNET_DEF_FLOW_DURATION As Long = 13
Public Const POLICY_ERRV_SUBNET_DEF_FLOW_RATE As Long = 21
Public Const POLICY_ERRV_SUBNET_DEF_PEAK_RATE As Long = 29
Public Const POLICY_ERRV_SUBNET_DEF_SUM_FLOW_RATE As Long = 37
Public Const POLICY_ERRV_SUBNET_DEF_SUM_PEAK_RATE As Long = 45
Public Const POLICY_ERRV_SUBNET_GRP_FLOW_COUNT As Long = 6
Public Const POLICY_ERRV_SUBNET_GRP_FLOW_DURATION As Long = 14
Public Const POLICY_ERRV_SUBNET_GRP_FLOW_RATE As Long = 22
Public Const POLICY_ERRV_SUBNET_GRP_PEAK_RATE As Long = 30
Public Const POLICY_ERRV_SUBNET_GRP_SUM_FLOW_RATE As Long = 38
Public Const POLICY_ERRV_SUBNET_GRP_SUM_PEAK_RATE As Long = 46
Public Const POLICY_ERRV_SUBNET_UNAUTH_USER_FLOW_COUNT As Long = 8
Public Const POLICY_ERRV_SUBNET_UNAUTH_USER_FLOW_DURATION As Long = 16
Public Const POLICY_ERRV_SUBNET_UNAUTH_USER_FLOW_RATE As Long = 24
Public Const POLICY_ERRV_SUBNET_UNAUTH_USER_PEAK_RATE As Long = 32
Public Const POLICY_ERRV_SUBNET_UNAUTH_USER_SUM_FLOW_RATE As Long = 40
Public Const POLICY_ERRV_SUBNET_UNAUTH_USER_SUM_PEAK_RATE As Long = 48
Public Const POLICY_ERRV_SUBNET_USER_FLOW_COUNT As Long = 7
Public Const POLICY_ERRV_SUBNET_USER_FLOW_DURATION As Long = 15
Public Const POLICY_ERRV_SUBNET_USER_FLOW_RATE As Long = 23
Public Const POLICY_ERRV_SUBNET_USER_PEAK_RATE As Long = 31
Public Const POLICY_ERRV_SUBNET_USER_SUM_FLOW_RATE As Long = 39
Public Const POLICY_ERRV_SUBNET_USER_SUM_PEAK_RATE As Long = 47
Public Const POLICY_ERRV_UNKNOWN As Long = 0
Public Const POLICY_ERRV_UNKNOWN_USER As Long = 49
Public Const POLICY_ERRV_UNSUPPORTED_CREDENTIAL_TYPE As Long = 2
Public Const POLICY_ERRV_USER_CHANGED As Long = 54
Public Const POLICY_GET_PRIVATE_INFORMATION As Long = &H4&
Public Const POLICY_KERBEROS_VALIDATE_CLIENT As Long = &H80
Public Const POLICY_LOCATOR_SUB_TYPE_ASCII_DN As Long = 1
Public Const POLICY_LOCATOR_SUB_TYPE_ASCII_DN_ENC As Long = 3
Public Const POLICY_LOCATOR_SUB_TYPE_UNICODE_DN As Long = 2
Public Const POLICY_LOCATOR_SUB_TYPE_UNICODE_DN_ENC As Long = 4
Public Const POLICY_LOOKUP_NAMES As Long = &H800&
Public Const POLICY_NOTIFICATION As Long = &H1000&
Public Const POLICY_QOS_ALLOW_LOCAL_ROOT_CERT_STORE As Long = &H20
Public Const POLICY_QOS_DHCP_SERVER_ALLOWED As Long = &H80
Public Const POLICY_QOS_INBOUND_CONFIDENTIALITY As Long = &H10
Public Const POLICY_QOS_INBOUND_INTEGRITY As Long = &H8
Public Const POLICY_QOS_OUTBOUND_CONFIDENTIALITY As Long = &H4
Public Const POLICY_QOS_OUTBOUND_INTEGRITY As Long = &H2
Public Const POLICY_QOS_RAS_SERVER_ALLOWED As Long = &H40
Public Const POLICY_QOS_SCHANNEL_REQUIRED As Long = &H1
Public Const POLICY_SERVER_ADMIN As Long = &H400&
Public Const POLICY_SET_AUDIT_REQUIREMENTS As Long = &H100&
Public Const POLICY_SET_DEFAULT_QUOTA_LIMITS As Long = &H80&
Public Const POLICY_TRUST_ADMIN As Long = &H8&
Public Const POLICY_VIEW_AUDIT_INFORMATION As Long = &H2&
Public Const POLICY_VIEW_LOCAL_INFORMATION As Long = &H1&
Public Const READ_CONTROL As Long = &H20000
Public Const STANDARD_RIGHTS_EXECUTE As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const STANDARD_RIGHTS_WRITE As Long = (READ_CONTROL)

Public Const POLICY_EXECUTE As Long = (STANDARD_RIGHTS_EXECUTE Or POLICY_VIEW_LOCAL_INFORMATION Or POLICY_LOOKUP_NAMES)
Public Const POLICY_READ As Long = (STANDARD_RIGHTS_READ Or POLICY_VIEW_AUDIT_INFORMATION Or POLICY_GET_PRIVATE_INFORMATION)
Public Const POLICY_WRITE As Long = (STANDARD_RIGHTS_WRITE Or POLICY_TRUST_ADMIN Or POLICY_CREATE_ACCOUNT Or POLICY_CREATE_SECRET Or POLICY_CREATE_PRIVILEGE Or POLICY_SET_DEFAULT_QUOTA_LIMITS Or POLICY_SET_AUDIT_REQUIREMENTS Or POLICY_AUDIT_LOG_ADMIN Or POLICY_SERVER_ADMIN)

Public Const POLICY_AUDIT_EVENT_MASK As Long = (POLICY_AUDIT_EVENT_SUCCESS Or POLICY_AUDIT_EVENT_FAILURE Or POLICY_AUDIT_EVENT_UNCHANGED Or POLICY_AUDIT_EVENT_NONE)
Public Const POLICY_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or POLICY_VIEW_LOCAL_INFORMATION Or POLICY_VIEW_AUDIT_INFORMATION Or POLICY_GET_PRIVATE_INFORMATION Or POLICY_TRUST_ADMIN Or POLICY_CREATE_ACCOUNT Or POLICY_CREATE_SECRET Or POLICY_CREATE_PRIVILEGE Or POLICY_SET_DEFAULT_QUOTA_LIMITS Or POLICY_SET_AUDIT_REQUIREMENTS Or POLICY_AUDIT_LOG_ADMIN Or POLICY_SERVER_ADMIN Or POLICY_LOOKUP_NAMES)

Public Const SE_ASSIGNPRIMARYTOKEN_NAME As String = "SeAssignPrimaryTokenPrivilege"
Public Const SE_AUDIT_NAME As String = "SeAuditPrivilege"
Public Const SE_BACKUP_NAME As String = "SeBackupPrivilege"
Public Const SE_BATCH_LOGON_NAME As String = "SeBatchLogonRight"
Public Const SE_CHANGE_NOTIFY_NAME As String = "SeChangeNotifyPrivilege"
Public Const SE_CREATE_PAGEFILE_NAME As String = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT_NAME As String = "SeCreatePermanentPrivilege"
Public Const SE_CREATE_TOKEN_NAME As String = "SeCreateTokenPrivilege"
Public Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
Public Const SE_DENY_BATCH_LOGON_NAME As String = "SeDenyBatchLogonRight"
Public Const SE_DENY_INTERACTIVE_LOGON_NAME As String = "SeDenyInteractiveLogonRight"
Public Const SE_DENY_NETWORK_LOGON_NAME As String = "SeDenyNetworkLogonRight"
Public Const SE_DENY_SERVICE_LOGON_NAME As String = "SeDenyServiceLogonRight"
Public Const SE_ENABLE_DELEGATION_NAME As String = "SeEnableDelegationPrivilege"
Public Const SE_INC_BASE_PRIORITY_NAME As String = "SeIncreaseBasePriorityPrivilege"
Public Const SE_INCREASE_QUOTA_NAME As String = "SeIncreaseQuotaPrivilege"
Public Const SE_INTERACTIVE_LOGON_NAME As String = "SeInteractiveLogonRight"
Public Const SE_LOAD_DRIVER_NAME As String = "SeLoadDriverPrivilege"
Public Const SE_LOCK_MEMORY_NAME As String = "SeLockMemoryPrivilege"
Public Const SE_MACHINE_ACCOUNT_NAME As String = "SeMachineAccountPrivilege"
Public Const SE_NETWORK_LOGON_NAME As String = "SeNetworkLogonRight"
Public Const SE_PROF_SINGLE_PROCESS_NAME As String = "SeProfileSingleProcessPrivilege"
Public Const SE_REMOTE_SHUTDOWN_NAME As String = "SeRemoteShutdownPrivilege"
Public Const SE_RESTORE_NAME As String = "SeRestorePrivilege"
Public Const SE_SECURITY_NAME As String = "SeSecurityPrivilege"
Public Const SE_SERVICE_LOGON_NAME As String = "SeServiceLogonRight"
Public Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"
Public Const SE_SYNC_AGENT_NAME As String = "SeSyncAgentPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT_NAME As String = "SeSystemEnvironmentPrivilege"
Public Const SE_SYSTEM_PROFILE_NAME As String = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME_NAME As String = "SeSystemtimePrivilege"
Public Const SE_TAKE_OWNERSHIP_NAME As String = "SeTakeOwnershipPrivilege"
Public Const SE_TCB_NAME As String = "SeTcbPrivilege"
Public Const SE_UNDOCK_NAME As String = "SeUndockPrivilege"
Public Const SE_UNSOLICITED_INPUT_NAME As String = "SeUnsolicitedInputPrivilege"

Public Const STATUS_SUCCESS As Long = &H0
Public Const STATUS_NO_MORE_ENTRIES As Long = &H8000001A
Public Const ERROR_MR_MID_NOT_FOUND As Long = 317&

Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Public Declare Function LsaNtStatusToWinError Lib "advapi32.dll" (ByVal Status As Long) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    lpSource As Long, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Any _
    ) As Long
'---------------------------------------------------------------------------------------------

Public Const CP_ACP = 0

Public Type WSTR
    data As Integer
End Type

Public Type LSA_UNICODE_STRING
    Length As Integer
    MaximumLength As Integer
     Buffer As String
End Type

Public Type LSA_OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As LSA_UNICODE_STRING
    Attributes As Long
    SecurityDescriptor As Long ' Points to type SECURITY_DESCRIPTOR
    SecurityQualityOfService As Long ' Points to type
                                     ' SECURITY_QUALITY_OF_SERVICE
End Type

Public Type lsaArray
    lsaData(4000) As Byte
End Type

Public Type ReferenceDomainName
    nameData(128) As Byte
End Type

Public Type psid
    sidData(228) As Byte
End Type

Public Type pSidArray
    sidData(228) As Long
End Type

Public Declare Function LsaOpenPolicy Lib "advapi32.dll" _
   (SystemName As LSA_UNICODE_STRING, ObjectAttributes As _
   LSA_OBJECT_ATTRIBUTES, ByVal DesiredAccess As Long, _
   PolicyHandle As Long) As Long
Public Declare Function LsaClose Lib "advapi32.dll" _
   (ByVal PolicyHandle As Long) As Long

Public Declare Function LsaAddAccountRights Lib "advapi32.dll" _
   (ByVal PolicyHandle As Long, AccountSid As psid, userRights As _
   LSA_UNICODE_STRING, ByVal CountOfRights As Long) As Long
Public Declare Function LsaRemoveAccountRights Lib "advapi32.dll" _
   (ByVal PolicyHandle As Long, AccountSid As psid, ByVal AllRights _
   As Byte, userRights As LSA_UNICODE_STRING, ByVal CountOfRights _
   As Long) As Long

Public Declare Function LookupAccountName Lib "advapi32.dll" Alias _
   "LookupAccountNameA" (ByVal lpSystemName As String, ByVal _
   lpAccountName As String, Sid As psid, cbSid As Long, _
   ReferencedDomainName As ReferenceDomainName, _
   cbReferencedDomainName As Long, peUse As Long) As Long

Public Declare Function LsaEnumerateAccountsWithUserRight Lib _
   "advapi32.dll" (ByVal PolicyHandle As Long, userRights As _
   LSA_UNICODE_STRING, EnumerationBuffer As Long, CountOfSIDs As _
   Long) As Long
Public Declare Function LsaEnumerateAccountRights Lib "advapi32.dll" _
   (ByVal PolicyHandle As Long, AccountSid As psid, EnumerationBuffer _
   As Long, CountOfSIDs As Long) As Long

Public Declare Function LookupAccountSid Lib "advapi32.dll" Alias _
   "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As _
   Long, Name As ReferenceDomainName, cbName As Long, _
   ReferencedDomainName As ReferenceDomainName, _
   cbReferencedDomainName As Long, peUse As Long) As Long

Public Declare Function LsaFreeMemory Lib "advapi32.dll" (ByVal _
   lpBuffer As Long) As Long

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal _
   hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As _
   pSidArray, ByVal nSize As Long, lpNumberOfBytesWritten As Long) _
   As Long
Public Declare Function ReadProcessMemory2 Lib "kernel32" Alias _
   "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress _
   As Any, lpBuffer As lsaArray, ByVal nSize As Long, _
   lpNumberOfBytesWritten As Long) As Long

Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal _
   CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As _
   String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, _
   ByVal cchWideChar As Long) As Long
