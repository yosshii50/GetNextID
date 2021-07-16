VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.ListBox List1 
      Height          =   4560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    List1.Width = Me.ScaleWidth
    List1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Load()
    
    Dim WrkIdx As Single
    Dim WrkID As String
    
    List1.Clear
    
    For WrkIdx = 0 To 1000
        WrkID = GetNextID(WrkID)
        Call List1.AddItem(WrkID & vbTab & WrkIdx)
    Next
    
End Sub

'����ID�擾
Private Function GetNextID(NowNo As String) As String
    
    Dim RetCD As String
    
    '�ŉ����擾
    Dim LastChar As String
    LastChar = Right(NowNo, 1)
    
    '��ʌ��擾
    Dim UpStr As String
    If NowNo = "" Then
        UpStr = ""
    Else
        UpStr = Left(NowNo, Len(NowNo) - 1)
    End If
    
    Select Case LastChar
    Case "0" To "8", "A" To "Y"
        '���̋L���擾
        LastChar = Chr(Asc(LastChar) + 1)
    Case "9" '[9]�̎���[A]���g�p
        LastChar = "A"
    Case "Z"
        '�Ō�܂ōs�����̂Ŏ��̌���
        UpStr = GetNextID(UpStr)
        LastChar = "0"
    Case Else
        '�s���ȏꍇ�A0����X�^�[�g
        LastChar = "0"
    End Select
    
    GetNextID = UpStr & LastChar
End Function
