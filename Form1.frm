VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4800
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   240
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   240
      Top             =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
If App.PrevInstance Then MsgBox "�����Ѿ��ں�̨����", vbInformation, "��ʾ": End
End Sub

Private Sub Timer1_Timer()
If PingIP("180.97.33.107") = True Then
Label1 = "��������"
Else:
Label1 = "�����쳣"
mciSendString "close OpenFile", 0&, 0, 0    '�ر�����
mciSendString "open """ & App.Path & "\mp3\err.mp3"" alias OpenFile type MPEGVideo", 0&, 0, 0  'ָ�������ļ�,Ϊmp3��ʽ
mciSendString "play OpenFile", 0&, 0, 0     '��������
Sleep (4000)
mciSendString "close OpenFile", 0&, 0, 0    '�ر�����
mciSendString "open """ & App.Path & "\mp3\xf.mp3"" alias OpenFile type MPEGVideo", 0&, 0, 0  'ָ�������ļ�,Ϊmp3��ʽ
mciSendString "play OpenFile", 0&, 0, 0     '��������
�޸�����
'MsgBox "�����쳣", vbCritical, "����"
Timer1.Enabled = False
End If
End Sub

Private Sub �޸�����()
Shell "cmd /c ipconfig /release"
Shell "cmd /c ipconfig /renew"
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
If PingIP("180.97.33.107") = True Then
mciSendString "close OpenFile", 0&, 0, 0    '�ر�����
mciSendString "open """ & App.Path & "\mp3\cg.mp3"" alias OpenFile type MPEGVideo", 0&, 0, 0  'ָ�������ļ�,Ϊmp3��ʽ
mciSendString "play OpenFile", 0&, 0, 0     '��������
'MsgBox "�����޸��ɹ�", vbInformation, "��ʾ"
Sleep (4000)
Timer2.Enabled = False
Timer1.Enabled = True
Else:
'MsgBox "�����޸�ʧ��", vbCritical, "����"
mciSendString "close OpenFile", 0&, 0, 0    '�ر�����
mciSendString "open """ & App.Path & "\mp3\sb.mp3"" alias OpenFile type MPEGVideo", 0&, 0, 0  'ָ�������ļ�,Ϊmp3��ʽ
mciSendString "play OpenFile", 0&, 0, 0     '��������
Sleep (4000)
Timer2.Enabled = False
Timer1.Enabled = True
End If
End Sub
