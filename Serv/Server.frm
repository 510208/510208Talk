VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   Caption         =   "TCP_Server"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   9240
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8760
      Top             =   0
   End
   Begin VB.TextBox txtSend 
      Height          =   1935
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Server.frx":0000
      Top             =   480
      Width           =   6375
   End
   Begin VB.TextBox txtReceive 
      Height          =   1935
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Server.frx":0006
      Top             =   2520
      Width           =   6375
   End
   Begin MSWinsockLib.Winsock Winsock_Server 
      Left            =   8760
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "接收文字"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "傳送文字"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblState 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Winsock TCP 模式

Private Sub Form_Load()
Winsock_Server.Protocol = sckTCPProtocol
Winsock_Server.LocalPort = 6000         '本機連接埠
Winsock_Server.Listen                   '監聽
'=====================================================================
'以下為單一專案測試時，同時開啟2表單
Me.Top = 500
Me.Left = Screen.Width / 5
'Frm_tcpClient.Show                      '開啟用戶端介面

'Frm_tcpClient.Top = Frm_tcpServer.Top + Frm_tcpServer.Height
'Frm_tcpClient.Left = Frm_tcpServer.Left
'=====================================================================

lblState = "等待連線……"
txtReceive.Text = Empty
txtSend.Enabled = False                 '等待連線
End Sub
'=====================================================================

Private Sub Timer1_Timer()
If Winsock_Server.State = sckListening Then
    lblState = "等待連線……"
    txtSend.Enabled = False                 '等待連線
   
Else
    RemoteIP = Winsock_Server.RemoteHostIP
    lblState = RemoteIP & " 已連線……"
    txtSend.Enabled = True                 '已連線
End If
If Winsock_Server.State = sckClosing Then
    Winsock_Server.Close                    '關閉，重新等待連線
    Winsock_Server.Listen
End If

End Sub

'=====================================================================

Private Sub txtSend_Change()

If Winsock_Server.State = sckListening Then     '監聽中：無用戶端連線
    Call Winsock_Server_Error(vbError, "用戶端未連線!", vbError, Source, HelpFile, HelpContext, False)
    Exit Sub
End If
Winsock_Server.SendData txtSend.Text
End Sub

'=====================================================================
Private Sub Winsock_Server_ConnectionRequest(ByVal requestID As Long)
If Winsock_Server.State <> sckClosed Then Winsock_Server.Close

Winsock_Server.Accept requestID         '允許遠端連線

End Sub


'=====================================================================

Private Sub Winsock_Server_DataArrival(ByVal bytesTotal As Long)
Dim inData As String
Winsock_Server.GetData inData, vbString
txtReceive.Text = inData
End Sub

'=====================================================================

Private Sub Winsock_Server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, _
            ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
s = ""
s = s & "錯誤碼：" & Number & vbNewLine
s = s & "訊息：" & Description & vbNewLine
s = s & "Scode：" & Scode & vbNewLine
s = s & "錯誤來源：" & Source & vbNewLine
MsgBox s
End Sub

