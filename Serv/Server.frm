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
   StartUpPosition =   3  '�t�ιw�]��
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
      Caption         =   "������r"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "�ǰe��r"
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
'Winsock TCP �Ҧ�

Private Sub Form_Load()
Winsock_Server.Protocol = sckTCPProtocol
Winsock_Server.LocalPort = 6000         '�����s����
Winsock_Server.Listen                   '��ť
'=====================================================================
'�H�U����@�M�״��ծɡA�P�ɶ}��2���
Me.Top = 500
Me.Left = Screen.Width / 5
'Frm_tcpClient.Show                      '�}�ҥΤ�ݤ���

'Frm_tcpClient.Top = Frm_tcpServer.Top + Frm_tcpServer.Height
'Frm_tcpClient.Left = Frm_tcpServer.Left
'=====================================================================

lblState = "���ݳs�u�K�K"
txtReceive.Text = Empty
txtSend.Enabled = False                 '���ݳs�u
End Sub
'=====================================================================

Private Sub Timer1_Timer()
If Winsock_Server.State = sckListening Then
    lblState = "���ݳs�u�K�K"
    txtSend.Enabled = False                 '���ݳs�u
   
Else
    RemoteIP = Winsock_Server.RemoteHostIP
    lblState = RemoteIP & " �w�s�u�K�K"
    txtSend.Enabled = True                 '�w�s�u
End If
If Winsock_Server.State = sckClosing Then
    Winsock_Server.Close                    '�����A���s���ݳs�u
    Winsock_Server.Listen
End If

End Sub

'=====================================================================

Private Sub txtSend_Change()

If Winsock_Server.State = sckListening Then     '��ť���G�L�Τ�ݳs�u
    Call Winsock_Server_Error(vbError, "�Τ�ݥ��s�u!", vbError, Source, HelpFile, HelpContext, False)
    Exit Sub
End If
Winsock_Server.SendData txtSend.Text
End Sub

'=====================================================================
Private Sub Winsock_Server_ConnectionRequest(ByVal requestID As Long)
If Winsock_Server.State <> sckClosed Then Winsock_Server.Close

Winsock_Server.Accept requestID         '���\���ݳs�u

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
s = s & "���~�X�G" & Number & vbNewLine
s = s & "�T���G" & Description & vbNewLine
s = s & "Scode�G" & Scode & vbNewLine
s = s & "���~�ӷ��G" & Source & vbNewLine
MsgBox s
End Sub

