VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00D27800&
   BorderStyle     =   0  'None
   Caption         =   "MaxXSoft Examer"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9375
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picLbl 
      BackColor       =   &H00D27800&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   2415
      TabIndex        =   12
      Top             =   6840
      Width           =   2415
      Begin VB.Label lblShow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "��¼ 0, ���ڴ��� 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   2370
      End
   End
   Begin VB.Timer tmrSPic 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6255
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   720
      Width           =   9375
      Begin VB.TextBox txtCnsl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   4560
         Width           =   6975
      End
      Begin VB.ListBox lstPCs 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   6150
         Left            =   0
         TabIndex        =   11
         Top             =   -15
         Width           =   2415
      End
      Begin VB.PictureBox picPg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   0
         Left            =   2400
         ScaleHeight     =   4575
         ScaleWidth      =   6975
         TabIndex        =   18
         Top             =   0
         Width           =   6975
         Begin ExamerSvr.ucBtn btnExit 
            Height          =   615
            Left            =   600
            TabIndex        =   22
            Top             =   3360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   1085
            Caption         =   "�˳� (�����ﲻ�Ƽ�Ŷ)"
            Default         =   0   'False
         End
         Begin ExamerSvr.ucBtn btnStart 
            Height          =   615
            Left            =   600
            TabIndex        =   21
            Top             =   1680
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   1085
            Caption         =   "���ٿ�ʼһ������"
            Default         =   -1  'True
         End
         Begin ExamerSvr.ucBtn btnDelAll 
            Height          =   615
            Left            =   600
            TabIndex        =   35
            Top             =   2520
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   1085
            Caption         =   "�������е��û�����"
            Default         =   0   'False
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "i"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   14.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   360
            Index           =   5
            Left            =   4320
            TabIndex        =   60
            Top             =   1800
            Width           =   285
         End
         Begin VB.Line linBrd 
            BorderColor     =   &H00D27800&
            Index           =   0
            X1              =   6960
            X2              =   6960
            Y1              =   0
            Y2              =   4560
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ǿ��������ҳ�������ѡ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   600
            TabIndex        =   20
            Top             =   960
            Width           =   3600
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��� (Hello world!)"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   18
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   465
            Index           =   3
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Width           =   3150
         End
      End
      Begin VB.PictureBox picPg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   2
         Left            =   2400
         ScaleHeight     =   4575
         ScaleWidth      =   6975
         TabIndex        =   23
         Top             =   0
         Width           =   6975
         Begin VB.PictureBox picListBg 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5F5F5&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2295
            Left            =   600
            ScaleHeight     =   2295
            ScaleWidth      =   5895
            TabIndex        =   29
            Top             =   1800
            Width           =   5895
            Begin ExamerSvr.ucDataList dlInfo 
               Height          =   2295
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Visible         =   0   'False
               Width           =   5895
               _ExtentX        =   10186
               _ExtentY        =   4048
            End
            Begin VB.Label lblShow 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "[����ɾ��]"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   21.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   570
               Index           =   11
               Left            =   0
               TabIndex        =   31
               Top             =   840
               Width           =   5880
            End
         End
         Begin ExamerSvr.ucBtn btnCopy 
            Height          =   375
            Left            =   3840
            TabIndex        =   28
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "���� IP"
            Default         =   0   'False
         End
         Begin ExamerSvr.ucBtn btnDel 
            Height          =   375
            Left            =   5280
            TabIndex        =   34
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "ɾ���û�"
            Default         =   0   'False
         End
         Begin VB.Label lblRfsh 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ˢ����Ϣ"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   5520
            TabIndex        =   32
            Top             =   1440
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lblShow 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "�������ͣ�ڡ�i����[...]��������ȡ��ϸ��Ϣ"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   10
            Left            =   600
            TabIndex        =   30
            Top             =   4200
            Width           =   5880
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϣ��(δ�����ܷ�)"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   9
            Left            =   600
            TabIndex        =   27
            Top             =   1440
            Width           =   2550
         End
         Begin VB.Label lblIP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "192.168.xx.xxx"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1920
            TabIndex        =   26
            Top             =   960
            Width           =   1590
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP ��ַ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   8
            Left            =   600
            TabIndex        =   25
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�û� XXX"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   18
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   465
            Index           =   7
            Left            =   360
            TabIndex        =   24
            Top             =   240
            Width           =   1500
         End
         Begin VB.Line linBrd 
            BorderColor     =   &H00D27800&
            Index           =   2
            X1              =   6960
            X2              =   6960
            Y1              =   0
            Y2              =   4560
         End
      End
      Begin VB.PictureBox picPg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   1
         Left            =   2400
         ScaleHeight     =   4575
         ScaleWidth      =   6975
         TabIndex        =   15
         Top             =   0
         Width           =   6975
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����Ŀǰ..."
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   18
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   465
            Index           =   6
            Left            =   360
            TabIndex        =   17
            Top             =   240
            Width           =   1710
         End
         Begin VB.Label lblShow 
            BackStyle       =   0  'Transparent
            Caption         =   "��Ŷ������͵���ˣ���ҳ��ʩ���С���"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   795
            Index           =   5
            Left            =   600
            TabIndex        =   16
            Top             =   960
            Width           =   5685
         End
         Begin VB.Line linBrd 
            BorderColor     =   &H00D27800&
            Index           =   1
            X1              =   6960
            X2              =   6960
            Y1              =   0
            Y2              =   4560
         End
      End
      Begin VB.PictureBox picPg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   4
         Left            =   2400
         ScaleHeight     =   4575
         ScaleWidth      =   6975
         TabIndex        =   50
         Top             =   0
         Width           =   6975
         Begin VB.PictureBox picLogo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1440
            Left            =   1440
            Picture         =   "frmMain.frx":000C
            ScaleHeight     =   1440
            ScaleWidth      =   4320
            TabIndex        =   58
            Top             =   840
            Width           =   4320
         End
         Begin VB.CheckBox chkBkDr 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Backdoors."
            Height          =   375
            Left            =   2160
            TabIndex        =   59
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ȩ���� 2010-2014  MaxXSoft ��������"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   300
            Index           =   22
            Left            =   600
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   55
            Top             =   4080
            Width           =   4155
         End
         Begin VB.Label lblShow 
            BackStyle       =   0  'Transparent
            Caption         =   "�ر��л����Ȫһ�м�������Լ�������֧�ֿ����￪������"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   420
            Index           =   21
            Left            =   600
            TabIndex        =   54
            Top             =   3480
            Width           =   5670
         End
         Begin VB.Label lblShow 
            BackStyle       =   0  'Transparent
            Caption         =   "������ʵ�������Ҫ��MaxXSoft ����������һ�������������Ҳ�ɴ˵���������ͼ����ƣ�MaxXSoft Gsy."
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   705
            Index           =   20
            Left            =   600
            TabIndex        =   53
            Top             =   2880
            Width           =   5745
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MaxXSoft Examer �汾 x.x.x"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   19
            Left            =   600
            TabIndex        =   52
            Top             =   2400
            Width           =   2715
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ڿ�����"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   18
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   465
            Index           =   18
            Left            =   360
            TabIndex        =   51
            Top             =   240
            Width           =   1800
         End
         Begin VB.Line linBrd 
            BorderColor     =   &H00D27800&
            Index           =   4
            X1              =   6960
            X2              =   6960
            Y1              =   0
            Y2              =   4560
         End
      End
      Begin VB.PictureBox picPg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   3
         Left            =   2400
         ScaleHeight     =   4575
         ScaleWidth      =   6975
         TabIndex        =   36
         Top             =   0
         Width           =   6975
         Begin VB.CheckBox chkDtl 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5880
            TabIndex        =   57
            Top             =   2520
            Width           =   375
         End
         Begin VB.TextBox txtQlib 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3720
            TabIndex        =   45
            Top             =   3240
            Width           =   2895
         End
         Begin VB.TextBox txtMax 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   5880
            TabIndex        =   42
            Text            =   "0"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txtPort 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   5880
            TabIndex        =   41
            Text            =   "80"
            Top             =   1080
            Width           =   735
         End
         Begin ExamerSvr.ucBtn btnSave 
            Height          =   495
            Left            =   5160
            TabIndex        =   38
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "��������"
            Default         =   -1  'True
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "i"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   12.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   330
            Index           =   4
            Left            =   2880
            TabIndex        =   56
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "i"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   12.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   330
            Index           =   3
            Left            =   1680
            TabIndex        =   49
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "i"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   12.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   330
            Index           =   2
            Left            =   2640
            TabIndex        =   48
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "i"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   12.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   330
            Index           =   1
            Left            =   4200
            TabIndex        =   47
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "i"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   12.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   330
            Index           =   0
            Left            =   2400
            TabIndex        =   46
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ļ�"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   17
            Left            =   600
            TabIndex        =   44
            Top             =   3240
            Width           =   960
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʾ������ϸ��Ϣ"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   15
            Left            =   600
            TabIndex        =   43
            Top             =   2520
            Width           =   1920
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������� (��Ϊ""0""����������)"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   14
            Left            =   600
            TabIndex        =   40
            Top             =   1800
            Width           =   3450
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���Է������˿�"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   13
            Left            =   600
            TabIndex        =   39
            Top             =   1080
            Width           =   1680
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����������..."
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   18
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D27800&
            Height          =   465
            Index           =   12
            Left            =   360
            TabIndex        =   37
            Top             =   240
            Width           =   2430
         End
         Begin VB.Line linBrd 
            BorderColor     =   &H00D27800&
            Index           =   3
            X1              =   6960
            X2              =   6960
            Y1              =   0
            Y2              =   4560
         End
      End
      Begin VB.Label lblBg 
         BackColor       =   &H00D27800&
         Height          =   135
         Left            =   0
         TabIndex        =   14
         Top             =   6120
         Width           =   2415
      End
   End
   Begin VB.PictureBox picTlt 
      BackColor       =   &H00D27800&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9375
      TabIndex        =   3
      Top             =   0
      Width           =   9375
      Begin VB.Label lblBtn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblBtn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblBtn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblBtn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblBtn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҳ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   420
      End
      Begin VB.Image imgCtrl 
         Height          =   480
         Index           =   1
         Left            =   8400
         Picture         =   "frmMain.frx":10D4
         Stretch         =   -1  'True
         ToolTipText     =   "��С������"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgCtrl 
         Height          =   480
         Index           =   0
         Left            =   8880
         Picture         =   "frmMain.frx":1133
         Stretch         =   -1  'True
         ToolTipText     =   "���Ǹ����"
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MaxXSoft Examer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   9465
      End
   End
   Begin VB.Label lblShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������ʼ����..."
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   1
      Left            =   4320
      TabIndex        =   2
      Top             =   3720
      Width           =   3360
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome Text..."
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1050
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   9405
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Shdw As cShadow, bTip As Boolean

Sub ChangeCap(sTxt As String)
    Me.Caption = "[" & sTxt & "] MaxXSoft Examer"
    lblCap.Caption = Me.Caption
End Sub

Sub ConsoleAdd(sTxt As String, Optional lSound As Long = 0)
    On Error Resume Next
    Dim sTxtAdd As String
    If lSound = 0 Then
        sTxtAdd = " [" & Time & "] " & sTxt & vbCrLf
    Else
        MessageBeep lSound
        Dim sSnd() As String
        sSnd = Split("����,ѯ��,����,��ʾ", ",")
        sTxtAdd = " [" & Time & " " & sSnd(lSound / 16 - 1) & "] " & sTxt & vbCrLf
    End If
    With txtCnsl
        If Len(.Text & sTxtAdd) > 32765 Then .Text = ""
        .Text = .Text & sTxtAdd
        .SelStart = Len(.Text)
    End With
End Sub

Sub RefreshStatus(lRecords As Long, lPrcss As Long)
    lblShow(2).Caption = "��¼ " & lRecords & ", ���ڴ��� " & lPrcss
End Sub

Sub SetOrder(cCtrl As Control)
    cCtrl.ZOrder 0
    txtCnsl.ZOrder 0
    lstPCs.ZOrder 0
End Sub

Private Sub btnCopy_Click()
    Clipboard.Clear
    Clipboard.SetText lblIP.Caption
    ConsoleAdd "IP ��ַ�ѳɹ����Ƶ������壡"
End Sub

Private Sub btnDel_Click()
    If MsgBox("��ȷ��Ҫɾ������û���" & vbCrLf & "�����ｨ������ִ�д˲���ǰ��ȷ���û������Ѿ���ȫ����", 48 + vbOKCancel, "�ֻ���ô~") = vbCancel Then Exit Sub
    DelUser lblIP.Caption
    lblBtn_Click 2
    ConsoleAdd "������ɹ�ɾ���� IP " & lblIP.Caption & " ���û�����..."
End Sub

Private Sub btnDelAll_Click()
    If frmIcon.sckHtp(0).State <> 0 Then
        ConsoleAdd "��������ȫ���ǣ������ﲻ�����ڿ��Խ���ʱִ�д˲���", 48
        Exit Sub
    End If
    If ReadLib("UserData") = "" Then ConsoleAdd "������ʲô�û�������~": Exit Sub
    If MsgBox("��ȷ��Ҫɾ�������û����ݣ�", 48 + vbOKCancel, "�ֻ���ô~") = vbCancel Then Exit Sub
    SaveUser True
    ConsoleAdd "������ɹ�������û�����..."
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    If frmIcon.sckHtp(0).State <> 0 Then
        ConsoleAdd "��������ȫ���ǣ������ﲻ�����ڿ��Խ���ʱִ�д˲���", 48
        Exit Sub
    End If
    Dim sErrMsg As String
    sErrMsg = ""
    If Not IsNumeric(txtPort.Text) Or Not IsNumeric(txtMax.Text) Then
        sErrMsg = "����д��ȷ�Ķ˿� (��������) ��Ϣ"
    ElseIf (CLng(txtPort.Text) < 0 Or CLng(txtPort.Text) > 65535) Or (CLng(txtMax.Text) < 0 Or CLng(txtMax.Text) > 65535) Then
        sErrMsg = "�˿� (��������) ��Ϣ�е㲻��ѧ"
    End If
    If Dir(MyPath & txtQlib.Text) = "" Then sErrMsg = "δ�ҵ�����ļ�����˲��ļ����Ƿ���д��ȷ"
    If sErrMsg <> "" Then
        ConsoleAdd sErrMsg, 48
        Exit Sub
    End If
    SaveCon "Port", txtPort.Text
    SaveCon "MaxLink", txtMax.Text
    SaveCon "ShowDetails", CStr(chkDtl.Value)
    SaveCon "Qlib", txtQlib.Text
    ConsoleAdd "�����������������Ϣ~"
End Sub

Private Sub btnStart_Click()
    If btnStart.Caption = "��������" Then
        frmIcon.CloseServer
    Else
        Dim sItems() As String, i As Long
        sItems = Split("Port,MaxLink,Qlib,ShowDetails", ",")
        For i = 0 To UBound(sItems)
            If ReadCon(CStr(sItems(i))) = "" Then
                ConsoleAdd "�����ÿ�ȱŶ~��ȷ������ҳ�п��������Ϣ��д����", 48
                Exit Sub
            End If
        Next i
        If Dir(MyPath & ReadCon("Qlib")) = "" Then
            ConsoleAdd "������δ�ҵ�����ļ����������ļ���λ���Ƿ�������ȷ", 48
            Exit Sub
        End If
        If Not CheckQlib Then
            ConsoleAdd "��������Ϊ����ļ����ݳ����˴�����������ļ��Ƿ���Ϲ淶", 48
            Exit Sub
        End If
        frmIcon.InitServer CLng(ReadCon("Port"))
    End If
    lblBtn_Click 0
End Sub

Private Sub chkBkDr_KeyDown(KeyCode As Integer, Shift As Integer)
    If chkBkDr.Value <> 2 Then chkBkDr.Value = 2
End Sub

Private Sub dlInfo_MarkSet()
    lblRfsh_Click
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then End
    lblShow(0).Caption = GetWelText
    ChangeCap "��ʼ����..."
    RefreshStatus 0, 0
    picFrm.Left = Me.ScaleWidth
    picTlt.Top = -picTlt.Height
    lblBtn(0).Top = 400
    SetOrder picPg(0)
    setBorderColor lstPCs.hwnd, &HD27800
    setBorderColor txtCnsl.hwnd, &HD27800
    lblBg.Move 0, lstPCs.Top + lstPCs.Height, lstPCs.Width, picFrm.ScaleHeight - (lstPCs.Top + lstPCs.Height)
    picLbl.Move 0, picFrm.Top + picFrm.Height - lblBg.Height, lblBg.Width
    lblShow(2).Move 0, picLbl.Height, picLbl.ScaleWidth
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 180
    Set Shdw = New cShadow
    Set Tip = New CTooltip
    With Shdw
        .Color = vbBlack
        .Depth = 0
        .Transparency = 120
        .Shadow Me
    End With
    With Tip
        .Style = TTStandard
        .Icon = TTIconInfo
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmrSPic.Tag = "" Then Exit Sub
    On Error Resume Next
    Static ox!, oy!
    With Me
        If Button = 1 Then
            .Move .Left - ox + X, .Top - oy + Y
        Else
            ox = X
            oy = Y
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmIcon.CloseServer
    Unload frmIcon
    Tip.Destroy
    Set Shdw = Nothing
    Set Tip = Nothing
End Sub

Private Sub imgCtrl_Click(Index As Integer)
    If Index = 0 Then
        Unload Me
    Else
        Me.WindowState = 1
    End If
End Sub

Private Sub imgCtrl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgCtrl(Index)
        .Move IIf(Index = 0, Me.ScaleWidth - .Width, Me.ScaleWidth - .Width * 2) + 10, 10
    End With
End Sub

Private Sub imgCtrl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgCtrl(Index)
        .Move IIf(Index = 0, Me.ScaleWidth - .Width, Me.ScaleWidth - .Width * 2) - 10, 0
    End With
End Sub

Private Sub lblBtn_Click(Index As Integer)
    Dim i As Long
    For i = 0 To lblBtn.UBound
        lblBtn(i).Top = 480
    Next i
    lblBtn(Index).Top = 400
    picPg(Index).Move lstPCs.Width - 15, 0
    Select Case Index
        Case 0
            If frmIcon.sckHtp(0).State = 0 Then
                btnStart.Caption = "���ٿ�ʼһ������"
            Else
                btnStart.Caption = "��������"
            End If
        Case 2
            With lstPCs
                dlInfo.Visible = False
                lblShow(9).Caption = "������Ϣ��(δ�����ܷ�)"
                lblShow(10).Visible = False
                If .ListIndex <> -1 Then
                    lblShow(7).Caption = "�û� " & Replace(Mid(.Text, InStrRev(.Text, "[") + 1), "]", "")
                    lblIP.Caption = Replace(.Text, "[" & Replace(lblShow(7).Caption, "�û� ", "") & "]", "")
                    btnCopy.Visible = True
                    btnDel.Visible = True
                    lblRfsh.Visible = True
                    lblRfsh_Click
                Else
                    lblShow(7).Caption = "δѡ���û�"
                    lblIP.Caption = lblShow(11).Caption
                    btnCopy.Visible = False
                    btnDel.Visible = False
                    lblRfsh.Visible = False
                End If
            End With
        Case 3
            txtPort.Text = ReadCon("Port")
            txtMax.Text = ReadCon("MaxLink")
            chkDtl.Value = CInt(IIf(ReadCon("ShowDetails") = "", 0, ReadCon("ShowDetails")))         '��ֹ settings.xcfg �ڲ�����
            txtQlib.Text = ReadCon("Qlib")
        Case 4
            lblShow(19).Caption = "MaxXSoft Examer �汾 " & _
                App.Major & "." & App.Minor & "." & App.Revision & " " & Trim(Replace(App.FileDescription, App.ProductName, ""))
    End Select
    SetOrder picPg(Index)
End Sub

Private Sub lblCap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bTip Then
        Dim sTips() As String
        sTips = Split("��������˿����ｫ���ĸ��˿ڴ�������" & vbCrLf & "Ĭ��Ϊ 80," & _
            "�������ʱ�û��Է������ķ�����������趨ֵ��" & vbCrLf & "���ʽ����ܾ�," & _
            "���ÿ���ʱ����̨�Ƿ���ʾ�������ϸ��Ϣ," & _
            "�뽫����ļ����뿼�������ڵ�Ŀ¼��" & vbCrLf & "Ȼ�������������ļ��� (����չ��)��" & vbCrLf & "����ļ��淶��ο������ĵ�," & _
            "��ҳ��������������Ϣ���� XCfg �ļ���׼��" & vbCrLf & "���� XCfgEditor �༭��" & vbCrLf & "�����ļ��洢�ڡ�settings.xcfg��," & _
            "�û����ݽ��ڽ�������ʱ���棬�ڴ��ڼ佫һֱ�ݴ��ڳ����ڡ�" & vbCrLf & "��ע�⼰ʱ����", ",")
        With Tip
            .Title = "��Ҫ�Ķ���"
            .TipText = sTips(Index)
            .Create picPg(IIf(Index < 5, 3, 0)).hwnd
        End With
        bTip = True
    End If
End Sub

Private Sub lblRfsh_Click()
    dlInfo.Visible = dlInfo.ReloadData(lblIP.Caption)
    lblShow(10).Visible = dlInfo.Visible
    If dlInfo.Visible Then
        Dim lMkd As Long
        lMkd = dlInfo.GetMarked(lblIP.Caption)
        If lMkd = -1 Then
            lblShow(9).Caption = "������Ϣ��(δ�����ܷ�)"
        Else
            lblShow(9).Caption = "������Ϣ��(�ܷ� " & CStr(lMkd) & " ��)"
        End If
    Else
        lblShow(9).Caption = "������Ϣ��(δ�����ܷ�)"
    End If
End Sub

Private Sub lblShow_Click(Index As Integer)
    If Index = 22 Then Shell "rundll32.exe url.dll,FileProtocolHandler http://maxxsoft.net/", vbNormalFocus
End Sub

Private Sub lblShow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lstPCs_Click()
    lblBtn_Click 2
End Sub

Private Sub picLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub picLogo_DblClick()
    picLogo.Move 1440, 840
End Sub

Private Sub picLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static ox!, oy!
    With picLogo
        If Button = 1 And Shift = 2 Then
            .Move .Left - ox + X, .Top - oy + Y
        Else
            ox = X
            oy = Y
        End If
    End With
End Sub

Private Sub picPg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 3 And bTip Then
        Tip.Destroy
        bTip = False
    End If
End Sub

Private Sub picTlt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub tmrSPic_Timer()
    Select Case tmrSPic.Tag
        Case ""
            Shdw.Depth = GetMoveNum(15, CSng(Shdw.Depth), 1, 1)
            Me.Top = (Screen.Height - Me.Height) / 2 - Shdw.Depth * 12
            If Shdw.Depth = 15 Then
                tmrSPic.Tag = "0"
                InitApp                      '��ʼ�����򣬰����� Winsock �ؼ��Ƿ���ע����ж�
                Sleep 1000
            End If
        Case Else
            picTlt.Top = picTlt.Top + GetMoveNum(0, picTlt.Top, 7)
            lblShow(2).Top = lblShow(2).Top + GetMoveNum(0, lblShow(2).Top, 7)
            picFrm.Left = picFrm.Left + GetMoveNum(0, picFrm.Left, 7)
            If GetMoveNum(0, picFrm.Left, 7) = 0 Then
                ChangeCap "׼������"
                ConsoleAdd "������ȴ�����"
                frmIcon.DoCheckUpdate
                tmrSPic.Enabled = False
            End If
    End Select
End Sub
