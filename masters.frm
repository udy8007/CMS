VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form masters 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTERS INFORMATION"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "masters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRecord 
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   7440
      TabIndex        =   94
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   39
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1440
      TabIndex        =   40
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   2640
      TabIndex        =   41
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3840
      TabIndex        =   42
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   6240
      TabIndex        =   44
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecord 
      Cancel          =   -1  'True
      Caption         =   "Re&fresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   5040
      TabIndex        =   43
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   77
      Top             =   7560
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   240
      TabIndex        =   45
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   -2147483627
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Subject"
      TabPicture(0)   =   "masters.frx":27A2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(1)=   "Subject"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Course"
      TabPicture(1)   =   "masters.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Cource"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Staff"
      TabPicture(2)   =   "masters.frx":27DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Staff"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Holiday"
      TabPicture(3)   =   "masters.frx":27F6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Vendor"
      TabPicture(4)   =   "masters.frx":2812
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Item"
      TabPicture(5)   =   "masters.frx":282E
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Frame4"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin MSAdodcLib.Adodc Staff 
         Height          =   330
         Left            =   -69480
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "ty02"
         Password        =   "ty02"
         RecordSource    =   "SELECT * FROM STAFFMASTER"
         Caption         =   "Attendance"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Subject 
         Height          =   570
         Left            =   -73920
         Top             =   5160
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1005
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "ty02"
         Password        =   "ty02"
         RecordSource    =   "SELECT * FROM SubjectMaster"
         Caption         =   "Attendance"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Frame Frame6 
         Caption         =   "Subject Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74400
         TabIndex        =   51
         Top             =   840
         Width           =   7335
         Begin VB.TextBox txtSubjectCode 
            DataField       =   "SUBJECTCODE"
            DataSource      =   "Subject"
            Height          =   315
            Left            =   3240
            TabIndex        =   1
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            DataField       =   "MAXMARKS"
            DataSource      =   "Subject"
            Height          =   375
            Left            =   3240
            TabIndex        =   6
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox Text2 
            DataField       =   "MINMARKS"
            DataSource      =   "Subject"
            Height          =   375
            Left            =   3240
            TabIndex        =   5
            Top             =   2640
            Width           =   855
         End
         Begin VB.TextBox Text1 
            DataField       =   "REMARK"
            DataSource      =   "Subject"
            Height          =   975
            Left            =   3240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   3600
            Width           =   3855
         End
         Begin VB.TextBox unittxt 
            DataField       =   "NOOFUNIT"
            DataSource      =   "Subject"
            Height          =   375
            Left            =   3240
            TabIndex        =   4
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox subnmtxt 
            DataField       =   "SUBJECTNAME"
            DataSource      =   "Subject"
            Height          =   375
            Left            =   3240
            TabIndex        =   2
            Top             =   1560
            Width           =   3855
         End
         Begin MSDataListLib.DataCombo CourceCode 
            Bindings        =   "masters.frx":284A
            DataField       =   "COURCECODE"
            DataSource      =   "Subject"
            Height          =   315
            Left            =   3240
            TabIndex        =   0
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "COURCECODE"
            BoundColumn     =   "COURCECODE"
            Text            =   ""
            Object.DataMember      =   "CourceInfo"
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Maximum Mark:"
            Height          =   210
            Left            =   1560
            TabIndex        =   81
            Top             =   3120
            Width           =   1320
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Minimum Mark:"
            Height          =   210
            Left            =   1560
            TabIndex        =   80
            Top             =   2640
            Width           =   1290
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Remark"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   79
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "COURSE CODE"
            Height          =   210
            Left            =   1560
            TabIndex        =   76
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "NO. OF UNITS"
            Height          =   210
            Index           =   0
            Left            =   1560
            TabIndex        =   54
            Top             =   2160
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "SUBJECT NAME"
            Height          =   210
            Index           =   0
            Left            =   1560
            TabIndex        =   53
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "SUJECT CODE"
            Height          =   210
            Index           =   0
            Left            =   1560
            TabIndex        =   52
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000A&
         Caption         =   "Staff's Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -74520
         TabIndex        =   50
         Top             =   600
         Width           =   7335
         Begin VB.TextBox Text6 
            DataField       =   "REMARK"
            DataSource      =   "Staff"
            Height          =   855
            Left            =   4440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Text            =   "masters.frx":287F
            Top             =   3960
            Width           =   2775
         End
         Begin VB.TextBox Text5 
            DataField       =   "QUALIFICATION"
            DataSource      =   "Staff"
            Height          =   1095
            Left            =   4680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   1800
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Bindings        =   "masters.frx":2884
            DataField       =   "JOINDATE"
            DataSource      =   "Staff"
            Height          =   375
            Left            =   3720
            TabIndex        =   20
            Top             =   3120
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            _Version        =   393216
            Format          =   24379392
            CurrentDate     =   37673
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Bindings        =   "masters.frx":28A6
            DataField       =   "BIRTHDATE"
            DataSource      =   "Staff"
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   3120
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            _Version        =   393216
            Format          =   24379392
            CurrentDate     =   37673
         End
         Begin VB.ComboBox cboDesignation 
            DataField       =   "DESIGNATION"
            DataSource      =   "Staff"
            Height          =   330
            ItemData        =   "masters.frx":28C9
            Left            =   360
            List            =   "masters.frx":28DF
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   3960
            Width           =   1935
         End
         Begin VB.TextBox stasaltxt 
            DataField       =   "SALARY"
            DataSource      =   "Staff"
            Height          =   315
            Left            =   2520
            TabIndex        =   22
            Top             =   3960
            Width           =   1815
         End
         Begin VB.TextBox staphtxt 
            DataField       =   "PHONE"
            DataSource      =   "Staff"
            Height          =   285
            Left            =   2760
            TabIndex        =   17
            Text            =   "0"
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox staaddtxt 
            DataField       =   "ADDRESS"
            DataSource      =   "Staff"
            Height          =   975
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox txtEmpCode 
            DataField       =   "EMPCODE"
            DataSource      =   "Staff"
            Height          =   285
            Left            =   3720
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtLastName 
            DataField       =   "LASTNAME"
            DataSource      =   "Staff"
            Height          =   285
            Left            =   360
            TabIndex        =   13
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox stasnmtxt 
            DataField       =   "PARENTNAME"
            DataSource      =   "Staff"
            Height          =   285
            Left            =   4920
            TabIndex        =   15
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox stafnmtxt 
            DataField       =   "FIRSTNAME"
            DataSource      =   "Staff"
            Height          =   285
            Left            =   2400
            TabIndex        =   14
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Remark"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4440
            TabIndex        =   87
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label36 
            Caption         =   "Qualification"
            Height          =   255
            Left            =   4680
            TabIndex        =   86
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label35 
            Caption         =   "JoinDate"
            Height          =   255
            Left            =   3720
            TabIndex        =   85
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label34 
            Caption         =   "BirthDate"
            Height          =   255
            Left            =   360
            TabIndex        =   84
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Designation"
            Height          =   210
            Left            =   360
            TabIndex        =   83
            Top             =   3720
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "SALARY"
            Height          =   255
            Left            =   2520
            TabIndex        =   64
            Top             =   3720
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "PHONE   "
            Height          =   255
            Left            =   2760
            TabIndex        =   63
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "ADDRESS :"
            Height          =   255
            Left            =   360
            TabIndex        =   62
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "SurName"
            Height          =   255
            Left            =   360
            TabIndex        =   61
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "Second Name"
            Height          =   255
            Left            =   4920
            TabIndex        =   60
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "First Name"
            Height          =   255
            Left            =   2400
            TabIndex        =   59
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Employee Code"
            Height          =   255
            Left            =   2400
            TabIndex        =   58
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Item Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   480
         TabIndex        =   49
         Top             =   960
         Width           =   7335
         Begin VB.TextBox Text11 
            DataField       =   "REMARK"
            DataSource      =   "ItemMaster"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   3360
            Width           =   3375
         End
         Begin VB.TextBox Text10 
            DataField       =   "QTYONHAND"
            DataSource      =   "ItemMaster"
            Height          =   315
            Left            =   1920
            TabIndex        =   35
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox Text9 
            DataField       =   "MAXQTY"
            DataSource      =   "ItemMaster"
            Height          =   315
            Left            =   5160
            TabIndex        =   37
            Top             =   2760
            Width           =   1095
         End
         Begin VB.TextBox Text8 
            DataField       =   "MINQTY"
            DataSource      =   "ItemMaster"
            Height          =   315
            Left            =   3600
            TabIndex        =   36
            Top             =   2760
            Width           =   1095
         End
         Begin VB.TextBox itpricetxt 
            DataField       =   "PRICE"
            DataSource      =   "ItemMaster"
            Height          =   315
            Left            =   3360
            TabIndex        =   34
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtItemName 
            DataField       =   "ITEMNAME"
            DataSource      =   "ItemMaster"
            Height          =   315
            Left            =   3360
            TabIndex        =   3
            Top             =   1080
            Width           =   3855
         End
         Begin VB.TextBox txtItemCode 
            DataField       =   "ITEMCODE"
            DataSource      =   "ItemMaster"
            Height          =   375
            Left            =   3360
            TabIndex        =   33
            Top             =   480
            Width           =   1215
         End
         Begin MSAdodcLib.Adodc ItemMaster 
            Height          =   495
            Left            =   1920
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
            OLEDBString     =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   "ty02"
            Password        =   "ty02"
            RecordSource    =   "SELECT * FROM ITEMMASTER"
            Caption         =   "Attendance"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label41 
            BackColor       =   &H00000000&
            Caption         =   "Quantity"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   93
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Remark"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   92
            Top             =   3360
            Width           =   855
         End
         Begin VB.Shape Shape1 
            Height          =   975
            Left            =   1800
            Top             =   2280
            Width           =   4695
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "On Hand"
            DataSource      =   "ItemMaster"
            Height          =   210
            Left            =   1920
            TabIndex        =   91
            Top             =   2520
            Width           =   675
         End
         Begin VB.Label Label18 
            Caption         =   "Max"
            DataSource      =   "ItemMaster"
            Height          =   255
            Left            =   5160
            TabIndex        =   90
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Min"
            DataSource      =   "ItemMaster"
            Height          =   255
            Left            =   3600
            TabIndex        =   89
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "PRICE"
            DataSource      =   "ItemMaster"
            Height          =   255
            Left            =   1800
            TabIndex        =   67
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "ITEM NAME"
            DataSource      =   "ItemMaster"
            Height          =   375
            Left            =   1800
            TabIndex        =   66
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "ITEM CODE"
            DataSource      =   "ItemMaster"
            Height          =   255
            Left            =   1920
            TabIndex        =   65
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Vendor Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   -74400
         TabIndex        =   48
         Top             =   840
         Width           =   7335
         Begin VB.TextBox Text7 
            DataField       =   "REMARK"
            DataSource      =   "Vendor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   3240
            Width           =   3615
         End
         Begin VB.TextBox txtVendorCode 
            DataField       =   "VENDORCODE"
            DataSource      =   "Vendor"
            Height          =   285
            Left            =   2160
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox vphtxt 
            DataField       =   "PHONE"
            DataSource      =   "Vendor"
            Height          =   285
            Left            =   2160
            TabIndex        =   30
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox vaddtxt 
            DataField       =   "ADDRESS"
            DataSource      =   "Vendor"
            Height          =   855
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   29
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox vsnmtxt 
            DataField       =   "EMAIL"
            DataSource      =   "Vendor"
            Height          =   285
            Left            =   2160
            TabIndex        =   31
            Top             =   2760
            Width           =   1815
         End
         Begin VB.TextBox txtVendorName 
            DataField       =   "VENDORNAME"
            DataSource      =   "Vendor"
            Height          =   285
            Left            =   2160
            TabIndex        =   28
            Top             =   840
            Width           =   1575
         End
         Begin MSAdodcLib.Adodc Vendor 
            Height          =   495
            Left            =   840
            Top             =   3990
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   873
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
            OLEDBString     =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   "ty02"
            Password        =   "ty02"
            RecordSource    =   "SELECT * FROM VendorMaster"
            Caption         =   "Vendor"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Remark"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   495
            TabIndex        =   88
            Top             =   3240
            Width           =   1500
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "PHONE"
            Height          =   255
            Left            =   495
            TabIndex        =   73
            Top             =   2280
            Width           =   1500
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "ADDRESS"
            Height          =   255
            Left            =   500
            TabIndex        =   72
            Top             =   1200
            Width           =   1500
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "EMail"
            Height          =   255
            Left            =   500
            TabIndex        =   71
            Top             =   2760
            Width           =   1500
         End
         Begin VB.Label Label21 
            Caption         =   "SECOND NAME"
            Height          =   255
            Left            =   3120
            TabIndex        =   70
            Top             =   3600
            Width           =   1695
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "NAME"
            Height          =   255
            Left            =   500
            TabIndex        =   69
            Top             =   840
            Width           =   1500
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "VENDOR CODE"
            Height          =   255
            Left            =   500
            TabIndex        =   68
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Holiday Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -74520
         TabIndex        =   47
         Top             =   840
         Width           =   7335
         Begin VB.TextBox txtRemark 
            DataField       =   "REMARK"
            DataSource      =   "Holiday"
            Height          =   1095
            Left            =   3240
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   26
            Top             =   1920
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker HolidayDate 
            DataField       =   "HOLIDAYDATE"
            DataSource      =   "Holiday"
            Height          =   435
            Left            =   3240
            TabIndex        =   24
            Top             =   600
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   767
            _Version        =   393216
            Format          =   24379392
            CurrentDate     =   37609
         End
         Begin VB.TextBox honmtxt 
            DataField       =   "HOLIDAY"
            DataSource      =   "Holiday"
            Height          =   405
            Left            =   3240
            TabIndex        =   25
            Top             =   1320
            Width           =   3855
         End
         Begin MSAdodcLib.Adodc Holiday 
            Height          =   495
            Left            =   5760
            Top             =   3600
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
            OLEDBString     =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   "ty02"
            Password        =   "ty02"
            RecordSource    =   "SELECT * FROM HolidayMaster"
            Caption         =   "Attendance"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Remark"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2280
            TabIndex        =   78
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "HOLIDAY NAME"
            Height          =   375
            Left            =   1560
            TabIndex        =   75
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "HOLIDAY DATE"
            Height          =   255
            Left            =   1560
            TabIndex        =   74
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cource Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -74400
         TabIndex        =   46
         Top             =   960
         Width           =   7335
         Begin VB.TextBox Text4 
            DataField       =   "REMARK"
            DataSource      =   "Cource"
            Height          =   855
            Left            =   2520
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   3000
            Width           =   4575
         End
         Begin VB.TextBox durtxt 
            DataField       =   "DURATION"
            DataSource      =   "Cource"
            Height          =   375
            Left            =   2520
            TabIndex        =   10
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox cournmtxt 
            DataField       =   "COURCENAME"
            DataSource      =   "Cource"
            Height          =   375
            Left            =   2520
            TabIndex        =   9
            Top             =   1680
            Width           =   4695
         End
         Begin VB.TextBox courcodetxt 
            DataField       =   "COURCECODE"
            DataSource      =   "Cource"
            Height          =   315
            Left            =   2520
            TabIndex        =   8
            Top             =   1020
            Width           =   2415
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Remark"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1200
            TabIndex        =   82
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "DURATION"
            Height          =   255
            Left            =   1200
            TabIndex        =   57
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "COURSE NAME"
            Height          =   375
            Left            =   1200
            TabIndex        =   56
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "COURSE CODE "
            Height          =   255
            Left            =   1200
            TabIndex        =   55
            Top             =   1080
            Width           =   1335
         End
      End
      Begin MSAdodcLib.Adodc Cource 
         Height          =   615
         Left            =   -71880
         Top             =   5160
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;Password=ty02;User ID=ty02;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "ty02"
         Password        =   "ty02"
         RecordSource    =   "SELECT * FROM COURCEMASTER"
         Caption         =   "Attendance"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "masters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Option Explicit
 Dim rs As ADODB.Recordset
 Dim CBOAREA As String
 
Private Sub cmdhome_Click()
    Unload Me
    College.Show
End Sub

Private Sub cmdRecord_Click(Index As Integer)
'On Error GoTo RecordError
    If Index = 0 Then
        PreviousRecord
    ElseIf Index = 1 Then
        AddRecord
    ElseIf Index = 2 Then
        UpdateRecord
    ElseIf Index = 3 Then
        DeleteRecord
    ElseIf Index = 4 Then
        RefreshRecord
    ElseIf Index = 5 Then
        NextRecord
    Else
        Unload Me
        College.Show
    End If
    Exit Sub
'RecordError:
'        MsgBox Err.Description, vbExclamation, "Masters Form Error"
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
End Sub


Public Sub PreviousRecord()
On Error GoTo PreviousError
    Select Case SSTab1.Tab
        Case 0
            With Subject.Recordset
                .MovePrevious
                If .BOF Then
                    MsgBox "First Record", vbCritical, "First Record"
                    .MoveFirst
                End If
            End With
        Case 1
            With Cource.Recordset
                .MovePrevious
                If .BOF Then
                    MsgBox "First Record", vbCritical, "First Record"
                    .MoveFirst
                End If
            End With
        Case 2
            With Staff.Recordset
                .MovePrevious
                If .BOF Then
                    MsgBox "First Record", vbCritical, "First Record"
                    .MoveFirst
                End If
            End With
          Case 3
            With Holiday.Recordset
                .MovePrevious
                If .BOF Then
                    MsgBox "First Record", vbCritical, "First Record"
                    .MoveFirst
                End If
            End With
        Case 4
            With Vendor.Recordset
                .MovePrevious
                If .BOF Then
                    MsgBox "First Record", vbCritical, "First Record"
                    .MoveFirst
                End If
            End With
        Case 5
            With ItemMaster.Recordset
                .MovePrevious
                If .BOF Then
                    MsgBox "First Record", vbCritical, "First Record"
                    .MoveFirst
                End If
            End With
    End Select
Exit Sub
PreviousError:
    MsgBox Err.Description, vbCritical, "Record Move Error"
End Sub

Public Sub AddRecord()
On Error GoTo AddError
    Dim md As Long
    Select Case SSTab1.Tab
        Case 0
            Subject.Recordset.AddNew
            CourceCode.SetFocus
        Case 1
            Cource.Recordset.AddNew
            CourceCode.SetFocus
        Case 2
            With Staff.Recordset
                .MoveLast
                md = .Fields(0) + 1
                .AddNew
                txtEmpCode.Text = md
                txtLastName.SetFocus
            End With
        Case 3
            Holiday.Recordset.AddNew
            HolidayDate.Value = Date
            HolidayDate.SetFocus
        Case 4
              With Vendor.Recordset
                .MoveLast
                md = .Fields(0) + 1
                .AddNew
                txtVendorCode.Text = md
                txtVendorName.SetFocus
            End With
        Case 5
            With ItemMaster.Recordset
                .MoveLast
                md = .Fields(0) + 1
                .AddNew
                txtItemCode.Text = md
                txtItemName.SetFocus
            End With
    End Select
Exit Sub
AddError:
    MsgBox Err.Description, vbCritical, "Add Record Error"
End Sub

Public Sub UpdateRecord()
'On Error GoTo UpdateError
    Select Case SSTab1.Tab
        Case 0
            Subject.Recordset.Update
            
        Case 1
            If courcodetxt.Text = "" Or cournmtxt.Text = "" Or durtxt.Text = "" Or Text4.Text = "" Then
                MsgBox "Can Not Insert Null Value", vbCritical, "Invalid Value"
                Exit Sub
            Else
                Cource.Recordset.Update
            End If
            
        Case 2
            If txtLastName.Text = "" Or stafnmtxt.Text = "" Or stasaltxt.Text = "" Or staaddtxt.Text = "" Or Text5.Text = "" Or staphtxt.Text = "" Or cboDesignation.Text = "" Or stasaltxt.Text = "" Or Text6.Text = "" Then
                MsgBox "Can Not Insert Null Value", vbCritical, "Invalid Value"
                Exit Sub
            Else
                Staff.Recordset.Update
            End If
        Case 3
            Holiday.Recordset.Update
        Case 4
            Vendor.Recordset.Update
        Case 5
            ItemMaster.Recordset.Update
    End Select
    MsgBox "Record has been saved", vbInformation, "Save Record"
Exit Sub
'UpdateError:
'    MsgBox Err.Description, vbCritical, "Update Error"
End Sub

Public Sub DeleteRecord()
'On Error GoTo DeleteError
    If MsgBox("Delete Record", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Record") = vbNo Then Exit Sub
    Select Case SSTab1.Tab
        Case 0
            Subject.Recordset.Delete
        Case 1
            Cource.Recordset.Delete
        Case 2
            Staff.Recordset.Delete
        Case 3
            Holiday.Recordset.Delete
        Case 4
            Vendor.Recordset.Delete
        Case 5
            ItemMaster.Recordset.Delete
    End Select
Exit Sub
'DeleteError:
 '   MsgBox Err.Description, vbCritical, "Delete Error"
End Sub

Public Sub NextRecord()
On Error GoTo NextError
    Select Case SSTab1.Tab
        Case 0
            With Subject.Recordset
                .MoveNext
                If .EOF Then
                    MsgBox "Last Record", vbCritical, "Last Record"
                    .MoveLast
                End If
            End With
        Case 1
            With Cource.Recordset
                If .EOF Then
                    MsgBox "Last Record", vbCritical, "Last Record"
                    .MoveLast
                End If
            End If
            End With
        Case 2
            With Staff.Recordset
                .MoveNext
                If .EOF Then
                    MsgBox "Last Record", vbCritical, "Last Record"
                    .MoveLast
                End If
            End With
          Case 3
            With Holiday.Recordset
                .MoveNext
                If .EOF Then
                    MsgBox "Last Record", vbCritical, "Last Record"
                    .MoveLast
                End If
            End With
        Case 4
            With Vendor.Recordset
                .MoveNext
                If .EOF Then
                    MsgBox "Last Record", vbCritical, "Last Record"
                    .MoveLast
                End If
            End With
        Case 5
            With ItemMaster.Recordset
                .MoveNext
                If .EOF Then
                    MsgBox "Last Record", vbCritical, "Last Record"
                    .MoveLast
                End If
            End With
    End Select
Exit Sub
NextError:
    MsgBox Err.Description, vbCritical, "Record Move Error"
End Sub

Public Sub RefreshRecord()
On Error GoTo RefreshError
    Select Case SSTab1.Tab
        Case 0
            Subject.Recordset.CancelUpdate
        Case 1
            Cource.Recordset.CancelUpdate
        Case 2
            Staff.Recordset.CancelUpdate
        Case 3
            Holiday.Recordset.CancelUpdate
        Case 4
            Vendor.Recordset.CancelUpdate
        Case 5
            ItemMaster.Recordset.CancelUpdate
    End Select
Exit Sub
RefreshError:
    MsgBox Err.Description, vbCritical, "Refresh Record Error"
End Sub


Private Sub staphtxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub txtVENDORname_GotFocus()
    txtVendorName.SelStart = 0
    txtVendorName.SelLength = Len(txtVendorName.Text)
End Sub
Private Sub txtVENDORNAME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

Private Sub txtLASTname_GotFocus()
    txtLastName.SelStart = 0
    txtLastName.SelLength = Len(txtLastName.Text)
End Sub
Private Sub txtLASTname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

Private Sub STAFNMTXT_GotFocus()
    stafnmtxt.SelStart = 0
    stafnmtxt.SelLength = Len(stafnmtxt.Text)
End Sub
Private Sub STAFNMTXT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

Private Sub STASNMTXT_GotFocus()
    stasnmtxt.SelStart = 0
    stasnmtxt.SelLength = Len(stasnmtxt.Text)
End Sub
Private Sub STASNMTXT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 123) Or (KeyAscii >= 65 And KeyAscii <= 91) Then Else KeyAscii = 0
End Sub

Private Sub STASALtxt_GotFocus()
    stasaltxt.SelStart = 0
    stasaltxt.SelLength = Len(stasaltxt.Text)
End Sub
Private Sub STASALtxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
End Sub

Private Sub ITPRICEtxt_GotFocus()
    itpricetxt.SelStart = 0
    itpricetxt.SelLength = Len(itpricetxt.Text)
End Sub
'Private Sub ITPRICEtxt_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then CBOAREA.SetFocus
'
'    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Else KeyAscii = 0
'    End Sub


