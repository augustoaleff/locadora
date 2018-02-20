Attribute VB_Name = "Module1"
Option Explicit

'Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, _
 'ByVal lpInputName As String, ByVal lpOutputName As String) As Long

'Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

'Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

'Public Const MAX_PATH = 260

'Public DATA_eXP As Date

Public CN1 As ADODB.Connection
'Public CN1_Teste As ADODB.Connection
'Public CN1_MATEU As ADODB.Connection
'Public CN1_MEGA As ADODB.Connection
'Public CN1_ELETRO As ADODB.Connection
'Public CN1_NFE As ADODB.Connection
'Public CN1_CPAG As ADODB.Connection
'Public CN1_ENTREGAS_MTB As ADODB.Connection
'Public CN1_TEMP As ADODB.Connection
'Public CN1_MTB As ADODB.Connection

'Public MSG_PEDIDOS_ORCAMENTOS As String
'Public MSG_SUBST_TRIB1 As String
'Public MSG_SUBST_TRIB2 As String
'Public MSG_SUBST_TRIB3 As String
'Public MSG_SUBST_TRIB4 As String

'Public CN_NET As ADODB.Connection
'Public dbel As Database
'Public Clie As Recordset
'public CHQ As Recordset
Public reg As ADODB.Recordset
Public REG2 As ADODB.Recordset
Public REG3 As ADODB.Recordset
'Public REG3_NFE As ADODB.Recordset
Public reg4 As ADODB.Recordset
'Public REG4_NFE As ADODB.Recordset
'Public REG5_NFE As ADODB.Recordset
'Public REG_ENTREGAS_MTB As ADODB.Recordset
Public reg5 As ADODB.Recordset
Public reg6 As ADODB.Recordset
Public reg7 As ADODB.Recordset
Public reg8 As ADODB.Recordset
Public REG9 As ADODB.Recordset





Public LOJA As String

Public STR_DSN As String

Public USER_GRAU As Integer
Public USER_NAME As String

'Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function ReleaseCapture Lib "user32" () As Long
'Public Const WM_CAP_DRIVER_CONNECT As Long = 1034
'Public Const WM_CAP_DRIVER_DISCONNECT As Long = 1035
'Public Const WM_CAP_GRAB_FRAME As Long = 1084
'Public Const WM_CAP_EDIT_COPY As Long = 1054
'Public Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
'Public Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
'Public Const WM_CLOSE = &H10
'Public mCapHwnd As Long




