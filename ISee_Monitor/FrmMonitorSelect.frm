VERSION 5.00
Begin VB.Form FrmMonitorSelect 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "ISee Monitor Test"
   ClientHeight    =   7275
   ClientLeft      =   -150
   ClientTop       =   135
   ClientWidth     =   11760
   Icon            =   "FrmMonitorSelect.frx":0000
   LinkTopic       =   "FrmMonitorSelect"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11760
   Begin VB.Frame Frame2 
      Caption         =   "This Frame Using Chinese BIG-5 ����{����T"
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   6120
      Width           =   11535
      Begin VB.Label Label8 
         Caption         =   "V0.9.23"
         Height          =   255
         Left            =   10320
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "���n��|�b�睊���A���|���������@�ѡA�n����v : Free�I���n�鹪�y���ɻP�ƻs���j�l�κؤl�Ψ�L�q���W�A���G����L�H�C"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   9975
      End
      Begin VB.Label Label6 
         Caption         =   "�i�ܦ���F�z�ҧƱ�W�[���\��δ��չϧΨѧ@�̰Ѧ�"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6840
         TabIndex        =   20
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   "http://blog.yam.com/user/melix.html"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3840
         MousePointer    =   10  '���W��
         TabIndex        =   19
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "���@Blog :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LbAH 
         Caption         =   "��"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   17
         Top             =   360
         Width           =   255
      End
      Begin VB.Label LbAH 
         Caption         =   "�b"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.Label LbAH 
         Caption         =   "��"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
      Begin VB.Label LbAH 
         Caption         =   "��"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "�{���s�@ by :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���餤��"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdMonitor 
      Caption         =   "Monitor"
      Enabled         =   0   'False
      Height          =   600
      Index           =   4
      Left            =   10800
      TabIndex        =   10
      Top             =   1440
      Width           =   800
   End
   Begin VB.CommandButton CmdMonitor 
      Caption         =   "Monitor"
      Enabled         =   0   'False
      Height          =   600
      Index           =   3
      Left            =   10800
      TabIndex        =   9
      Top             =   2160
      Width           =   800
   End
   Begin VB.CommandButton CmdMonitor 
      Caption         =   "Monitor"
      Enabled         =   0   'False
      Height          =   600
      Index           =   2
      Left            =   10800
      TabIndex        =   8
      Top             =   2880
      Width           =   800
   End
   Begin VB.CommandButton CmdMonitor 
      Caption         =   "Monitor"
      Enabled         =   0   'False
      Height          =   600
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   3600
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9120
      TabIndex        =   3
      Text            =   "1"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdWhatMon 
      Caption         =   "View Desktop range"
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "English"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select display monitor"
      Height          =   5295
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   11655
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   6480
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4200
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Which you want view:"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Change Language"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmMonitorSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Monitors class, contains a collection of monitors
'and functions for multiple monitors
Dim cMonitors As clsMonitors

'Monitor class, contains information about a monitor
Dim cMonitor As clsMonitor

'Rectangle structure, for determining
'monitors at a given position
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Constants for the return value when finding a monitor
Const MONITOR_DEFAULTTONULL = &H0       'If the monitor is not found, return 0
Const MONITOR_DEFAULTTOPRIMARY = &H1    'If the monitor is not found, return the primary monitor
Const MONITOR_DEFAULTTONEAREST = &H2    'If the monitor is not found, return the nearest monitor

Public MntCount As Long
Public UserChooseMonitor As Long
Public ChooseMonitorCommand As Long
Public CM2Left As Long
Public CM2Top As Long
Public CM2Height As Long
Public CM2Width As Long
Dim FrmV(1 To 4) As FrmV1 'It is Create 4 Windows Form use on Multi Monitors
Dim ChLangu As Boolean

Private Sub CmdMntFix()
Dim MyStartX As Long, MyStartY As Long
Dim MyRangeTX As Long, MyRangeTY As Long, MyRangeBX As Long, MyRangeBY As Long
Dim MyDpX As Single, MyDpY As Single
MyStartX = 4760 'the virtual X0
MyStartY = 3600 'the virtual Y0
MyRangeTY = 1080 'the top of frame1
MyRangeBY = 4400 'the bottom of frame1
MyRangeTX = 100 ' the left of frame1
MyRangeBX = 9000 'the right of frame1

For ia = 1 To MntCount
cMonitors.Refresh
    Set cMonitor = cMonitors.Monitors.Item(ia)
    With cMonitor
        CmdMonitor(ia).Move (.Left + MyStartX), (.Top + MyStartY), .Width, .Height
    
    End With
Next

For ib = 1 To MntCount
cMonitors.Refresh
    Set cMonitor = cMonitors.Monitors.Item(ib)
    If CmdMonitor(ib).Top > MyRangeBY Then
    With cMonitor
        For ic = 1 To MntCount
        CmdMonitor(ic).Move CmdMonitor(ic).Left, (CmdMonitor(ic).Top - (MyStartY - MyRangeTY))
        Next
    End With
    End If

Next
End Sub


Private Sub CmdMntD_Click()

End Sub

Private Sub CmdMonitor_Click(Index As Integer)

ChooseMonitorCommand = Index

'********set the display monitor
cMonitors.Refresh
    'Get the reference to the monitor
    Set cMonitor = cMonitors.Monitors.Item(Index)
    With cMonitor
  '  MsgBox "Left start on: " & .Left & vbCrLf & _
  '      "Top start on: " & .Top & vbCrLf & _
  '      "Right end on: " & .Right & vbCrLf & _
  '      "Bottom end on: " & .Bottom & vbCrLf & _
  '      "Width: " & .Width & vbCrLf & _
  '      "Height: " & .Height, vbInformation
        
    CM2Left = .Left * 15
    CM2Top = .Top * 15
    CM2Height = .Height * 15
    CM2Width = .Width * 15
    End With

'If Option3.Value = True Then
'    Me.Hide
'    Exit Sub
'End If

'If Option2.Value = True Then
'    CM2Color = "&H" & HexColor(CM2RGB(0).Text) & HexColor(CM2RGB(1).Text) & HexColor(CM2RGB(2).Text)
'End If
        
'***********Create New Object FrmV to Show**********
Unload FrmV(Index)
FrmV(Index).Show
FrmV(Index).Caption = "popup:" & Index
'FrmV1.Show
'***************************************************

End Sub

Private Sub cmdWhatMon_Click()

    Dim lMonitor As Long

    cMonitors.Refresh
        
        'Get the reference to the monitor
        Set cMonitor = cMonitors.Monitors.Item(Val(Combo1.Text))
    
    With cMonitor
    If Langu = 1 Then
        MsgBox "���t���: " & .Left & vbCrLf & _
        "���I���: " & .Top & vbCrLf & _
        "�k�t���: " & .Right & vbCrLf & _
        "�������: " & .Bottom & vbCrLf & _
        "�e��: " & .Width & vbCrLf & _
        "����: " & .Height, vbInformation
    Else
        MsgBox "Left start on: " & .Left & vbCrLf & _
        "Top start on: " & .Top & vbCrLf & _
        "Right end on: " & .Right & vbCrLf & _
        "Bottom end on: " & .Bottom & vbCrLf & _
        "Width: " & .Width & vbCrLf & _
        "Height: " & .Height, vbInformation
    End If
    End With
End Sub

Private Sub Command1_Click()
ChLangu = True
Langu = 0
Unload Me
End Sub

Private Sub Command2_Click()
    'LoadPicIn
'Text1.Text = LoadPicIn

Dim i As Integer
Me.Height = 6465 + 1200
Langu = 1
Label1.Caption = "��ܤ@�ӧA�n�˵��d�򪺿ù�"
Label2.Caption = "�ܧ�y��"
cmdWhatMon.Caption = "�˵��ୱ�d��"
Frame1.Caption = "��ܤ@�ӱz�n���ժ��ù� , ���۹��z����x���]�w����m ! ���h�P�E�����տù��]�n������ض� ( �ХΤ�r�˵����W������ )"
For i = 1 To 4
    CmdMonitor(i).Caption = "�ù�"
Next
Command3.Caption = "(&C)�����{��"


End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.Height = 6465
    'Setup the classes
    Set cMonitor = New clsMonitor
    Set cMonitors = New clsMonitors
    Me.Width = 11850
    'Center the form
    cMonitors.CenterFormOnMonitor Me
    
    MntCount = cMonitors.Monitors.Count
    
    '*********************************************************
    For ld_i = 1 To MntCount
        Combo1.AddItem (ld_i) 'load how many monitors
        CmdMonitor(ld_i).Enabled = True 'Set how many Monitor Command button Enable
    Next
    '*********************************************************
    '******************Fix the CmdMonitor(s)******************
    Call CmdMntFix
    
    '***************Pre Set FrmV1 in FrmV(1 to 4)**************************
    Set FrmV(1) = New FrmV1
    Set FrmV(2) = New FrmV1
    Set FrmV(3) = New FrmV1
    Set FrmV(4) = New FrmV1
    '***************Pre Set The CM2 Value**************************
    cMonitors.Refresh
    Set cMonitor = cMonitors.Monitors.Item(1)
    With cMonitor
    CM2Left = .Left * 15
    CM2Top = .Top * 15
    CM2Height = .Height * 15
    CM2Width = .Width * 15
    End With
    '***********************************************************************
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'******Close All Form****
For Each FrmG In Forms
    Unload FrmG
Next
'************************

If ChLangu = True Then
    ChLangu = False
    FrmWC.Show
End If
    
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ShellExecute 0, "Open", "http://blog.yam.com/user/melix.html", "", "", 3
End If
End Sub

