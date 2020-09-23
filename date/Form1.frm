VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datafile"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   " Scrivi file "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   2400
      TabIndex        =   32
      Top             =   720
      Width           =   2055
      Begin VB.CommandButton cmdSet 
         Caption         =   "Scrivi dati"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   360
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtD4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtM4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtY4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtY5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtM5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtD5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtY6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtM6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   19
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtD6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Creato il: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   540
         TabIndex        =   40
         Top             =   750
         Width           =   150
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   1120
         TabIndex        =   39
         Top             =   750
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Modificato il: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1125
         TabIndex        =   37
         Top             =   1590
         Width           =   150
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   540
         TabIndex        =   36
         Top             =   1590
         Width           =   150
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1125
         TabIndex        =   35
         Top             =   2430
         Width           =   150
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   540
         TabIndex        =   34
         Top             =   2430
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo accesso il: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   1260
      End
   End
   Begin VB.Frame fra1 
      Caption         =   " Leggi file "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   2055
      Begin VB.CommandButton cmdGet 
         Caption         =   "Leggi dati"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtD3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtM3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtY3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtD2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtM2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtY2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtY1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtM1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtD1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblU1 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo accesso il: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   540
         TabIndex        =   30
         Top             =   2430
         Width           =   150
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1125
         TabIndex        =   29
         Top             =   2430
         Width           =   150
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   28
         Top             =   1590
         Width           =   150
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1125
         TabIndex        =   27
         Top             =   1590
         Width           =   150
      End
      Begin VB.Label lblmod1 
         AutoSize        =   -1  'True
         Caption         =   "Modificato il: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1120
         TabIndex        =   25
         Top             =   750
         Width           =   150
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   24
         Top             =   750
         Width           =   150
      End
      Begin VB.Label lblCre1 
         AutoSize        =   -1  'True
         Caption         =   "Creato il: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   705
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, _
    lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, _
    lpLastWriteTime As FILETIME) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" ( _
    ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" ( _
    lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32.dll" ( _
    lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" ( _
    lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long) As Long
Private Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, _
    lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, _
    lpLastWriteTime As FILETIME) As Long
Private Declare Sub GetSystemTimeAsFileTime Lib "kernel32.dll" ( _
    lpSystemTimeAsFileTime As FILETIME)
Private Declare Function SystemTimeToFileTime Lib "kernel32.dll" ( _
    lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long


Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const CREATE_ALWAYS = 2
Private Const CREATE_NEW = 1
Private Const OPEN_ALWAYS = 4
Private Const OPEN_EXISTING = 3
Private Const TRUNCATE_EXISTING = 5
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
    
Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Sub cmdGet_Click()

Dim hFile As Long  ' handle to the opened file
Dim ctime As FILETIME  ' receives time of creation
Dim atime As FILETIME  ' receives time of last access
Dim mtime As FILETIME  ' receives time of last modification
Dim thetime As SYSTEMTIME  ' used to manipulate the time
Dim retval As Long  ' return value
Dim giorno As Integer
Dim mese As Integer

hFile = CreateFile(txtPath.Text, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), _
    OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
If hFile = -1 Then
  Debug.Print "Could not open the file successfully -- aborting."
  Exit Sub  ' terminate the program
End If

' Next, get the creation, last-access, and last-modification times.
retval = GetFileTime(hFile, ctime, atime, mtime)
' Convert the creation time to the local time zone.
retval = FileTimeToLocalFileTime(ctime, ctime)
' Convert the FILETIME format to the SYSTEMTIME format.
retval = FileTimeToSystemTime(ctime, thetime)

txtD1 = Format(thetime.wDay, "00")
txtM1 = Format(thetime.wMonth, "00")
txtY1 = thetime.wYear

retval = FileTimeToLocalFileTime(mtime, mtime)
' Convert the FILETIME format to the SYSTEMTIME format.
retval = FileTimeToSystemTime(mtime, thetime)

txtD2 = Format(thetime.wDay, "00")
txtM2 = Format(thetime.wMonth, "00")
txtY2 = thetime.wYear

retval = FileTimeToLocalFileTime(atime, atime)
' Convert the FILETIME format to the SYSTEMTIME format.
retval = FileTimeToSystemTime(atime, thetime)

txtD3 = Format(thetime.wDay, "00")
txtM3 = Format(thetime.wMonth, "00")
txtY3 = thetime.wYear

' Close the file to free up resources.
retval = CloseHandle(hFile)
End Sub

Private Sub cmdSet_Click()
Dim hFile As Long  ' handle to the opened file
Dim ctime As FILETIME  ' the time of creation
Dim atime As FILETIME  ' the time of last access
Dim mtime As FILETIME  ' the time of last modification
Dim thetime1 As SYSTEMTIME
Dim thetime2 As SYSTEMTIME
Dim thetime3 As SYSTEMTIME
Dim retval As Long  ' return value

' First, open the file C:\MyApp\test.txt for both read-level and
' write-level access, since we need to do both.
hFile = CreateFile(txtPath.Text, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, _
    ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
If hFile = -1 Then
  Debug.Print "Could not open the file successfully -- aborting."
  Exit Sub  ' terminate the program
End If

' Next, get the creation, last-access, and last-modification times.
retval = GetFileTime(hFile, ctime, atime, mtime)

If txtD4.Text <> "" Then
    
    thetime1.wDay = CInt(txtD4.Text)
    thetime1.wMonth = CInt(txtM4.Text)
    thetime1.wYear = txtY4.Text
    thetime1.wHour = Hour(Now)
    thetime1.wMinute = Minute(Now)
    thetime1.wSecond = Second(Now)

    retval = SystemTimeToFileTime(thetime1, ctime)
    
    retval = LocalFileTimeToFileTime(ctime, ctime)
Else
    
    thetime1.wDay = Day(Now)
    thetime1.wMonth = Month(Now)
    thetime1.wYear = Year(Now)
    thetime1.wHour = Hour(Now)
    thetime1.wMinute = Minute(Now)
    thetime1.wSecond = Second(Now)
    
    retval = SystemTimeToFileTime(thetime1, ctime)
    
    retval = LocalFileTimeToFileTime(ctime, ctime)
    
End If

If txtD5.Text <> "" Then

    thetime2.wDay = CInt(txtD5.Text)
    thetime2.wMonth = CInt(txtM5.Text)
    thetime2.wYear = txtY5.Text
    thetime2.wHour = Hour(Now)
    thetime2.wMinute = Minute(Now)
    thetime2.wSecond = Second(Now)
    
    retval = SystemTimeToFileTime(thetime2, mtime)
    
    retval = LocalFileTimeToFileTime(mtime, mtime)

Else
    
    thetime2.wDay = Day(Now)
    thetime2.wMonth = Month(Now)
    thetime2.wYear = Year(Now)
    thetime2.wHour = Hour(Now)
    thetime2.wMinute = Minute(Now)
    thetime2.wSecond = Second(Now)
    
    retval = SystemTimeToFileTime(thetime2, mtime)
    
    retval = LocalFileTimeToFileTime(mtime, mtime)
    

End If

If txtD6.Text <> "" Then

    thetime3.wDay = CInt(txtD6.Text)
    thetime3.wMonth = CInt(txtM6.Text)
    thetime3.wYear = txtY6.Text
    thetime3.wHour = Hour(Now)
    thetime3.wMinute = Minute(Now)
    thetime3.wSecond = Second(Now)
    
    retval = SystemTimeToFileTime(thetime3, atime)
    
    retval = LocalFileTimeToFileTime(atime, atime)

Else

    thetime3.wDay = Day(Now)
    thetime3.wMonth = Month(Now)
    thetime3.wYear = Year(Now)
    thetime3.wHour = Hour(Now)
    thetime3.wMinute = Minute(Now)
    thetime3.wSecond = Second(Now)
    
    retval = SystemTimeToFileTime(thetime3, atime)
    
    retval = LocalFileTimeToFileTime(atime, atime)
    
End If
    
retval = SetFileTime(hFile, ctime, atime, mtime)

' Close the file to free up resources.
retval = CloseHandle(hFile)

End Sub

Private Sub cmdSfo_Click()

With cd1
    .Filter = "Tutti i file (*.*)|*.*"
    .ShowOpen
    If .FileName <> "" Then txtPath = .FileName
End With

End Sub
