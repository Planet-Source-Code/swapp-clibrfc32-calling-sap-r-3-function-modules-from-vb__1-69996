VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "External access to R/3 tables via RFC"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   45
      TabIndex        =   16
      Top             =   2115
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Country Key"
         Object.Width           =   2036
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Country Name"
         Object.Width           =   5424
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read table"
      Height          =   375
      Left            =   45
      TabIndex        =   15
      Top             =   4725
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "SAP Login"
      Height          =   1995
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4560
      Begin VB.CheckBox Check2 
         Caption         =   "ABAP Debugging"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1665
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Trace"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1395
         Width           =   1680
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3285
         PasswordChar    =   "*"
         TabIndex        =   12
         Text            =   "minisap"
         Top             =   630
         Width           =   1050
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   3285
         TabIndex        =   11
         Text            =   "BCUSER"
         Top             =   315
         Width           =   1050
      End
      Begin VB.TextBox txtLanguage 
         Height          =   285
         Left            =   3285
         TabIndex        =   10
         Text            =   "EN"
         Top             =   945
         Width           =   375
      End
      Begin VB.TextBox txtClient 
         Height          =   285
         Left            =   1260
         TabIndex        =   9
         Text            =   "000"
         Top             =   945
         Width           =   420
      End
      Begin VB.TextBox txtSysNo 
         Height          =   285
         Left            =   1260
         TabIndex        =   8
         Text            =   "00"
         Top             =   630
         Width           =   420
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1260
         TabIndex        =   7
         Text            =   "localhost"
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   195
         Index           =   5
         Left            =   2430
         TabIndex        =   6
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "User:"
         Height          =   195
         Index           =   4
         Left            =   2430
         TabIndex        =   5
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Language:"
         Height          =   195
         Index           =   3
         Left            =   2430
         TabIndex        =   4
         Top             =   990
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Client:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   990
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "System-No:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Host:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Tested with SAP Remote Function Call Library (librfc32.dll) 7000.0.95.5320
'and SAP Netweaver 7.0 SP12 (miniSAP)
'on Windows XP Professional SP2

'The (free) ABAP Trial Version of SAP NetWeaver 7.0 is available at http://www.sdn.sap.com

'Function module RFC_READ_TABLE
'External access to R/3 tables via RFC

'This is only an example how to use the Rfc Library with a standard function module
'For applications to read and export large tables from SAP I recommend to write your own
'function module, cause RFC_READ_TABLE is very slow

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
     ByRef Destination As Any, _
     ByRef Source As Any, _
     ByVal Length As Long)

Private Const RFC_OK As Long = &H0

'structures from the SAP Dictionary (SE11)
Private Type RFC_DB_OPT
    SO_TEXT As String * 72
End Type
Private Type RFC_DB_FLD
    FIELDNAME As String * 30
    DOFFSET As String * 6
    DDLENG As String * 6
    INTTYPE As String * 1
    AS4TEXT As String * 60
End Type
Private Type TAB512
    WA As String * 512
End Type

Private WithEvents LibRfc As CLibRfc32
Attribute LibRfc.VB_VarHelpID = -1

Private Sub Command1_Click()
  Dim iTable As Long
  Dim i As Long
  Dim OPTIONS As RFC_DB_OPT
  Dim FIELDS As RFC_DB_FLD
  Dim DATA As TAB512

    On Error GoTo ErrHandling

    If LibRfc.hRfc = 0 Then                     'already connected? (should be impossible)
        With LibRfc
            .Host = txtHost.Text
            .SystemNo = txtSysNo.Text
            .client = txtClient.Text
            .User = txtUsername.Text
            .Password = txtPassword.Text
            .Language = txtLanguage.Text
            .Trace = CBool(Check1.Value)        'will write a trace file in the application folder
            .Debugging = CBool(Check2.Value)    'start the ABAP Debugger
            .connect
        End With
    End If
    
    If LibRfc.hRfc = 0 Then Exit Sub            'check the rfc-handle
    
    With LibRfc
        .Initialize
        'Allocate space for 2 export parameters, 0 import parameters and 3 tables
        .AllocParamSpace 2, 0, 3
        '1st export parameter QUERY_TABLE (the table to read)
        .AddExportParam "QUERY_TABLE", "T005T"
        '2nd export parameter DELIMITER (sign for indicating field limits in DATA structure)
        .AddExportParam "DELIMITER", "|"
        '1st table OPTIONS (WHERE clauses)
        iTable = .AddNewTable("OPTIONS", Len(OPTIONS))
        'Add one entry into table OPTIONS
        OPTIONS.SO_TEXT = "SPRAS = '" & txtLanguage.Text & "'"      '<-- WHERE SPRAS = 'EN'
        .AddEntry iTable, VarPtr(OPTIONS), Len(OPTIONS)
        '2nd table FIELDS (fields to read)
        iTable = .AddNewTable("FIELDS", Len(FIELDS))
        'we want to read the country key
        FIELDS.FIELDNAME = "LAND1"                                  '<-- Country Key
        .AddEntry iTable, VarPtr(FIELDS), Len(FIELDS)
        'and the country name
        FIELDS.FIELDNAME = "LANDX"                                  '<-- Country Name
        .AddEntry iTable, VarPtr(FIELDS), Len(FIELDS)
        '3rd table DATA (data read out)
        iTable = .AddNewTable("DATA", Len(DATA))
        'Call the fm RFC_READ_TABLE
        If .CallReceiveExt("RFC_READ_TABLE") = RFC_OK Then
            If .Lines(iTable) > 0 Then
                For i = 1 To .Lines(iTable)
                    'Copy every line from the allocated space in the local DATA structure
                    CopyMemory DATA, ByVal CLng(.TableLine(iTable, i)), Len(DATA)
                    ListView1.ListItems.Add i, , Left$(DATA.WA, InStr(1, DATA.WA, "|") - 1)
                    ListView1.ListItems.Item(i).SubItems(1) = Trim$(Right$(DATA.WA, Len(DATA.WA) - InStr(1, DATA.WA, "|")))
                    'Debug.Print .LineAsString(iTable, i)
                Next i
            End If
        End If
    End With
    
ErrHandling:
    'Disconnect from the SAP system
    LibRfc.Disconnect
    'Free the allocated memory and clear all tables
    LibRfc.CleanUp
    
End Sub

Private Sub Form_Load()
    Set LibRfc = New CLibRfc32
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set LibRfc = Nothing
End Sub

Private Sub LibRfc_Error(ErrorNr As EErrorPosition, Message As String)
    MsgBox Message, vbExclamation, "Remote Function Call"
End Sub
