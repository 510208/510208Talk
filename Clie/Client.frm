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
   StartUpPosition =   3  '�t�ιw�]��
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
      Caption         =   "����"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "���_"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "�s�u"
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
      Caption         =   "�ϥΪ̦W��(&N)�G"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "������r(&T)�G"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "�ǰe��r(&S)�G"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�C���D��"
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
    MsgBox "���s�u!", vbCritical
    Exit Sub
End If

txtReceive.Text = Empty                  '�M�ű�����
txtSend.Text = Empty                     '�M�ű�����

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
'�]�w���s�u���D���W�١G
'�i�]��"127.0.0.1"
'�i�]�� "LocalHost"
'�i�]�� "�D���W��"
'�i�]�� "�D���ݤ�IP" �p "140.128.x.x"
'-------------------------------------------------------------------------
'Winsock_Client.RemoteHost = "LocalHost"     '�b����(LocalHost)����
'Winsock_Client.RemoteHost = "HSU_PC"         '���A�ݹq���W��

Winsock_Client.Protocol = sckTCPProtocol
Winsock_Client.RemoteHost = Combo1.Text
'--------------------------------------------------------------------------
Winsock_Client.RemotePort = 6000            '�]�w���A�ݩҶ}�񤧬ۦP�s����
Winsock_Client.Connect

End Sub

 

Private Sub btnEnd_Click()
Winsock_Client.Close
Unload Me
End Sub

'=====================================================================
Private Sub Form_Load()
Combo1.AddItem "127.1"                      '���A�ݬ�����
Combo1.AddItem "LocalHost"                      '���A�ݬ�����
Combo1.AddItem "HSU_PC"                     '���A�ݹq���W��
Combo1.AddItem "192.168.1.1"               '���A�ݰϺ�IP Address"
Combo1.AddItem "122.127.21.70"               '���A��Public IP Address"
Combo1.ListIndex = 0                         '�w�]��1�ӿﶵ
btnClose.Enabled = False
txtSend.Enabled = False
End Sub

'=====================================================================

Private Sub txtSend_Change()
'On Error GoTo error_Proc
txt = clieNameTxt.Text & "  " & Time & ":"
Winsock_Client.SendData txt & txtSend.Text
'(�ݼW�G���~�B�z)
'error_Proc:
    'MsgBox "���~�I�i��|���s�u�A���˹���~�T��" & vbNewLine & "�ԲӸ�T�G" & vbNewLine & Err.Description, vbCritical
End Sub

Private Sub Winsock_Client_DataArrival(ByVal bytesTotal As Long)
Dim inData As String
Winsock_Client.GetData inData, vbString
txtReceive.Text = inData

End Sub
