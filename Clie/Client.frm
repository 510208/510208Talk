VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Client 
   Caption         =   "TCP_Client"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   8130
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox clieNameTxt 
      Height          =   270
      Left            =   1560
      TabIndex        =   10
      Text            =   "Client1"
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtReceive 
      Height          =   1935
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2640
      Width           =   6735
   End
   Begin VB.TextBox txtSend 
      Height          =   1935
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   6735
   End
   Begin MSWinsockLib.Winsock Winsock_Client 
      Left            =   7680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton btnEnd 
      Caption         =   "結束"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "中斷"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "連線"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "使用者名稱(&N)："
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "接收文字(&T)："
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "傳送文字(&S)："
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "遊戲主機"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
If Winsock_Client.State = sckClosed Then
    MsgBox "未連線!", vbCritical
    Exit Sub
End If

txtReceive.Text = Empty                  '清空接收區
txtSend.Text = Empty                     '清空接收區

Winsock_Client.Close
btnClose.Enabled = False
btnConnect.Enabled = True
txtSend.Enabled = False
End Sub

Private Sub btnConnect_Click()
btnClose.Enabled = True
btnConnect.Enabled = False
txtSend.Enabled = True
'=========================================================================
'設定欲連線之主機名稱：
'可設為"127.0.0.1"
'可設為 "LocalHost"
'可設為 "主機名稱"
'可設為 "主機端之IP" 如 "140.128.x.x"
'-------------------------------------------------------------------------
'Winsock_Client.RemoteHost = "LocalHost"     '在本機(LocalHost)測試
'Winsock_Client.RemoteHost = "HSU_PC"         '伺服端電腦名稱

Winsock_Client.Protocol = sckTCPProtocol
Winsock_Client.RemoteHost = Combo1.Text
'--------------------------------------------------------------------------
Winsock_Client.RemotePort = 6000            '設定伺服端所開放之相同連接埠
Winsock_Client.Connect

End Sub

 

Private Sub btnEnd_Click()
Winsock_Client.Close
Unload Me
End Sub

'=====================================================================
Private Sub Form_Load()
Combo1.AddItem "127.1"                      '伺服端為本機
Combo1.AddItem "LocalHost"                      '伺服端為本機
Combo1.AddItem "HSU_PC"                     '伺服端電腦名稱
Combo1.AddItem "192.168.1.1"               '伺服端區網IP Address"
Combo1.AddItem "122.127.21.70"               '伺服端Public IP Address"
Combo1.ListIndex = 0                         '預設第1個選項
btnClose.Enabled = False
txtSend.Enabled = False
End Sub

'=====================================================================

Private Sub txtSend_Change()
'On Error GoTo error_Proc
txt = clieNameTxt.Text & "  " & Time & ":"
Winsock_Client.SendData txt & txtSend.Text
'(待增：錯誤處理)
'error_Proc:
    'MsgBox "錯誤！可能尚未連線，請檢察錯誤訊息" & vbNewLine & "詳細資訊：" & vbNewLine & Err.Description, vbCritical
End Sub

Private Sub Winsock_Client_DataArrival(ByVal bytesTotal As Long)
Dim inData As String
Winsock_Client.GetData inData, vbString
txtReceive.Text = inData

End Sub
