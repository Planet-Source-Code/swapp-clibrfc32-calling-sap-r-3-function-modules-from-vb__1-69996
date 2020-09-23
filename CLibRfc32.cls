VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLibRfc32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'For more informations about these functions read the Librfc SDK from SAP

' UDTs
Private Type RFC_ERROR_INFO_EX
    group   As Long
    Key     As String * 33
    Message As String * 513
End Type

Private Type RFC_ERROR_INFO
    Key     As String * 33
    Status  As String * 128
    Message As String * 256
    intstat As String * 128
End Type

'LIBRFC32
Private Declare Function RfcOpenEx Lib "librfc32.dll" (ByVal conString As String, ByRef RfcError As RFC_ERROR_INFO_EX) As Long
Attribute RfcOpenEx.VB_Description = "meldet die Anwendung am SAP-System an und gibt den einen Verbindungshandle zurück"
Private Declare Sub RfcClose Lib "librfc32.dll" (ByVal hRfc As Long)
Attribute RfcClose.VB_Description = "Schließt eine bestehende RFC-Verbindung"
Private Declare Function RfcLastError Lib "librfc32.dll" (ByRef RfcErrorInfo As RFC_ERROR_INFO) As Long
Attribute RfcLastError.VB_Description = "Zuletzt aufgetretenen Fehler lesen"
Private Declare Function RfcAllocParamSpace Lib "librfc32.dll" (ByVal numexp As Long, ByVal numimp As Long, ByVal numtab As Long) As Long
Attribute RfcAllocParamSpace.VB_Description = "Reserviert einen Speicherbereich, in den interne Tabellen (z.B. RFC_TABLE oder RFC_PARAMETER) mit RfcAddTable erstellt werden können"
Private Declare Function RfcFreeParamSpace Lib "librfc32.dll" (ByVal hSpace As Long) As Long
Attribute RfcFreeParamSpace.VB_Description = "Speicherbereich der Parameterübergabe wieder freigeben"
Private Declare Function RfcAddExportParam Lib "librfc32.dll" (ByVal hSpace As Long, ByVal parpos As Long, ByVal parname As String, ByVal parnamelen As Long, ByVal partype As Long, ByVal parlen As Long, ByRef par As Any) As Long
Attribute RfcAddExportParam.VB_Description = "Schreibt einen Exportparameter in den dafür reservierten Speicher"
Private Declare Function RfcAddImportParam Lib "librfc32.dll" (ByVal hSpace As Long, ByVal parpos As Long, ByVal parname As String, ByVal parnamelen As Long, ByVal partype As Long, ByVal parlen As Long, ByRef par As Any) As Long
Attribute RfcAddImportParam.VB_Description = "Schreibt einen Importparameter in den reservierten Speicherbereich"
Private Declare Function RfcAddTable Lib "librfc32.dll" (ByVal hSpace As Long, ByVal tabpos As Long, ByVal tabname As String, ByVal tabnamelen As Long, ByVal tabtype As Long, ByVal tablen As Long, ByVal tabhandle As Long) As Long
Attribute RfcAddTable.VB_Description = "Neue Tabelle in den Übergabebereich hinzufügen"
Private Declare Function RfcCallReceiveExt Lib "librfc32.dll" (ByVal hRfc As Long, ByVal hSpace As Long, ByVal funcname As String, ByVal Exception As String) As Long
Attribute RfcCallReceiveExt.VB_Description = "Aufruf des SAP-Funktionsbausteins"
Private Declare Function RfcDefineImportParam Lib "librfc32.dll" (ByVal hSpace As Long, ByVal ParamNo As Long, ByVal tabname As String, ByVal tabnamelen As Long, ByVal tabtype As Long, ByVal tablen As Long) As Long
Attribute RfcDefineImportParam.VB_Description = "Importparameter festlegen"
Private Declare Function RfcGetImportParam Lib "librfc32.dll" (ByVal hSpace As Long, ByVal ParamNo As Long, ByVal addr As Long) As Long
Attribute RfcGetImportParam.VB_Description = "Importparameter wieder auslesen"
Private Declare Function RfcLastErrorEx Lib "librfc32.dll" (ByRef ErrorInfo As RFC_ERROR_INFO_EX) As Long
Attribute RfcLastErrorEx.VB_Description = "Zuletzt aufgetretenen Fehler lesen"
Private Declare Function RfcIsValidHandle Lib "librfc32.dll" (ByVal hRfc As Long, ByRef ErrorInfo As RFC_ERROR_INFO_EX) As Long
Attribute RfcIsValidHandle.VB_Description = "Prüfen, ob der Rfc-Handle gültig ist"
Private Declare Sub RfcAbort Lib "librfc32.dll" (ByVal hRfc As Long, ByVal Message As String)
Private Declare Function RfcCallEx Lib "librfc32.dll" (ByVal hRfc As Long, ByVal funcname As String, ByVal exporting As Long, ByVal changing As Long, ByVal tables As Long) As Long
Private Declare Function RfcCallExt Lib "librfc32.dll" (ByVal hRfc As Long, ByVal hSpace As Long, ByVal funcname As String) As Long
Private Declare Function RfcCallReceiveEx Lib "librfc32.dll" (ByVal hRfc As Long, ByVal funcname As String, ByVal exporting As Long, ByVal changing As Long, ByVal tables As Long, ByVal Exception As String) As Long

Private Declare Function ItCreate Lib "librfc32.dll" (ByVal ItName As String, ByVal ItRecLen As Long, ByVal ItOccurs As Long, ByVal mem As Long) As Long
Attribute ItCreate.VB_Description = "Interne Tabelle anlegen"
Private Declare Function ItDelete Lib "librfc32.dll" (ByVal hIT As Long) As Long
Attribute ItDelete.VB_Description = "Interne Tabelle löschen"
Private Declare Function ItGetLine Lib "librfc32.dll" (ByVal hIT As Long, ByVal ItLine As Long) As Long
Attribute ItGetLine.VB_Description = "Zeiger auf Tabellenzeile ermitteln"
Private Declare Function ItInsLine Lib "librfc32.dll" (ByVal hIT As Long, ByVal ItLine As Long) As Long
Attribute ItInsLine.VB_Description = "Neue Zeile in interne Tabelle einfügen"
Private Declare Function ItAppLine Lib "librfc32.dll" (ByVal hIT As Long) As Long
Attribute ItAppLine.VB_Description = "Neue Zeile an interne Tabelle anhängen"
Private Declare Function ItDelLine Lib "librfc32.dll" (ByVal hIT As Long, ByVal ItLine As Long) As Long
Attribute ItDelLine.VB_Description = "Eintrag aus interner Tabelle löschen"
Private Declare Function ItGupLine Lib "librfc32.dll" (ByVal hIT As Long, ByVal ItLine As Long) As Long
Private Declare Function ItCpyLine Lib "librfc32.dll" (ByVal hIT As Long, ByVal ItLine As Long, ByVal dest As Long) As Long
Attribute ItCpyLine.VB_Description = "Eintrag einer internen Tabelle duplizieren"
Private Declare Function ItFree Lib "librfc32.dll" (ByVal hIT As Long) As Long
Attribute ItFree.VB_Description = "Speicherbereich der Tabelle freigeben"
Private Declare Function ItFill Lib "librfc32.dll" (ByVal hIT As Long) As Long
Attribute ItFill.VB_Description = "Anzahl der enthaltenen Tabellenzeilen ermitteln"
Private Declare Function ItLeng Lib "librfc32.dll" (ByVal hIT As Long) As Long
Attribute ItLeng.VB_Description = "Länge einer Tabellenzeile ermitteln"
Private Declare Function ItPutLine Lib "librfc32.dll" (ByVal hIT As Long, ByVal ItLine As Long, ByRef itContent As String) As String
Private Declare Function ItRegTable Lib "librfc32.dll" (ByVal hIT As Long, ByVal ufkt As Long, ByVal lpUdat As Long) As Long
Private Declare Function ItUserData Lib "librfc32.dll" (ByVal hIT As Long, ByRef lpUdat As Long) As Long

'From C header file
Private Const RFC_OK As Long = 0                    'OK
Private Const RFC_FAILURE As Long = 1               'error occured
Private Const RFC_EXCEPTION As Long = 2             'exception raised
Private Const RFC_SYS_EXCEPTION As Long = 3         'system exception raised, connection closed
Private Const RFC_CALL As Long = 4                  'call received
Private Const RFC_INTERNAL_COM As Long = 5          'internal communication, repeat
Private Const RFC_CLOSED As Long = 6                'connection closed by the other side
Private Const RFC_RETRY As Long = 7                 'no data yet (RfcListen only)
Private Const RFC_NO_TID As Long = 8                'no transaction ID available
Private Const RFC_EXECUTED As Long = 9              'function already executed
Private Const RFC_SYNCHRONIZE As Long = 10
Private Const RFC_MEMORY_INSUFFICIENT As Long = 11
Private Const RFC_VERSION_MISMATCH As Long = 12
Private Const RFC_NOT_FOUND As Long = 13
Private Const RFC_CALL_NOT_SUPPORTED As Long = 14
Private Const RFC_NOT_OWNER As Long = 15
Private Const RFC_NOT_INITIALIZED As Long = 16
Private Const RFC_SYSTEM_CALLED As Long = 17
Private Const RFC_INVALID_HANDLE As Long = 18
Private Const RFC_INVALID_PARAMETER As Long = 19
Private Const RFC_CANCELED As Long = 20
Private Const RFC_CONVERSION As Long = 21
Private Const RFC_INVALID_PROTOCOL As Long = 22
'KERNEL32
Private Declare Sub CopyMemoryA Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal ByteLen As Long)
Attribute CopyMemoryA.VB_Description = "Speicherbereiche kopieren"
Private Declare Function Incr Lib "kernel32.dll" Alias "InterlockedIncrement" (ByRef lpAddend As Long) As Long
Attribute Incr.VB_Description = "Wert inkrementieren (i++)"

Public Enum EErrorPosition
    EConnect
    EAllocateParamSpace
    EAddTable
    EAddExportParam
    EAddImportParam
    ECallReceiveExt
    EItCreate
    ECreateNewTable
    EAddNewTable
    EAddEntry
    EParameters
End Enum

Private Type TableProps
    Handle As Long
    name As String
End Type
Private Type ParamProps
    name As String
    Value As String
End Type

Private bLocked As Boolean

Private m_ConString As String
Private m_Exception As String
Private m_lpConString As Long

Private m_cntImport As Long
Private m_cntExport As Long
Private m_cntTables As Long

Private m_MaxImport As Long
Private m_MaxExport As Long
Private m_MaxTables As Long

Private m_hRfc As Long
Private m_hSpace As Long
Private m_Tables() As TableProps
Private m_Export() As ParamProps
Private m_Import() As ParamProps

Public Event Error(ByRef ErrorNr As EErrorPosition, ByRef Message As String)
Attribute Error.VB_Description = "Fehler"

Public Function connect() As Long
Attribute connect.VB_Description = "Am SAP-System anmelden"
  Dim RfcError As RFC_ERROR_INFO_EX
    If m_hRfc > 0 Then Exit Function
    m_hRfc = RfcOpenEx(m_ConString, RfcError)
    If m_hRfc = 0 Then
        RaiseEvent Error(EConnect, RfcError.Message)
        connect = 0
    Else
        connect = m_hRfc
    End If
End Function
Public Sub Disconnect()
Attribute Disconnect.VB_Description = "Verbindung zu SAP trennen"
    RfcClose m_hRfc
    m_hRfc = 0
End Sub

Public Property Get hRfc() As Long
Attribute hRfc.VB_Description = "Verbindungshandle zu SAP R/3"
Attribute hRfc.VB_UserMemId = 0
    hRfc = m_hRfc
End Property

Public Property Let Host(ByRef sParam As String)
Attribute Host.VB_Description = "SAP-Host"
    InsertLoginParam 7, sParam
End Property
Public Property Let User(ByRef sParam As String)
Attribute User.VB_Description = "SAP Benutzername"
    If Len(sParam) > 12 Then
        RaiseError EParameters, "Username"
    Else
        InsertLoginParam 65, sParam
    End If
End Property
Public Property Let Password(ByRef sParam As String)
Attribute Password.VB_Description = "SAP Paßwort"
    If Len(sParam) > 8 Then
        RaiseError EParameters, "Password"
    Else
        InsertLoginParam 85, sParam
    End If
End Property
Public Property Let Trace(ByRef bParam As Boolean)
Attribute Trace.VB_Description = "Trace-Modus einschalten"
    InsertLoginParam 108, CStr(Abs(bParam))
End Property
Public Property Let Debugging(ByRef bParam As Boolean)
Attribute Debugging.VB_Description = "Debug-Modus einschalten"
    InsertLoginParam 121, CStr(Abs(bParam))
End Property
Public Property Let SystemNo(ByRef sParam As String)
Attribute SystemNo.VB_Description = "SAP Systemnummer"
    InsertLoginParam 46, sParam
End Property
Public Property Let Language(ByRef sParam As String)
Attribute Language.VB_Description = "SAP Anmeldesprache"
    InsertLoginParam 99, sParam
End Property
Public Property Let client(ByRef sParam As String)
Attribute client.VB_Description = "SAP Mandant"
    InsertLoginParam 56, sParam
End Property

Public Property Get hSpace() As Long
    hSpace = m_hSpace
End Property

Public Property Get Exception() As String
Attribute Exception.VB_Description = "Beschreibung des aufgetretenen Fehlers (kommt aus LibRfc32)"
    Exception = Left$(m_Exception, InStr(1, m_Exception, vbNullChar))
End Property
Public Property Get ParamSpace() As Long
Attribute ParamSpace.VB_Description = "Zeiger auf den reservierten Speicher zur Parameterübergabe"
    ParamSpace = m_hSpace
End Property
Public Property Get hItab(ByVal Index As Long) As Long
Attribute hItab.VB_Description = "Handle der internen Tabelle"
    If m_cntTables = 0 Then Exit Property
    hItab = m_Tables(Index).Handle
End Property
Public Property Get TableHandleByName(ByRef sParam As String) As Long
Attribute TableHandleByName.VB_Description = "Rückgabe des Handles (hIT) anhand des Tabellennamens"
  Dim i As Long
    If m_cntTables = 0 Then Exit Property
    For i = 1 To UBound(m_Tables())
        If m_Tables(i).name = sParam Then
            TableHandleByName = m_Tables(i).Handle
            Exit For
        End If
    Next i
End Property
Public Property Get TableIndexByName(ByRef sParam As String) As Long
  Dim i As Long
    If m_cntTables = 0 Then Exit Property
    For i = 1 To UBound(m_Tables)
        If m_Tables(i).name = sParam Then
            TableIndexByName = i
            Exit For
        End If
    Next i
End Property
Public Property Get TableName(ByVal TableIndex As Long) As String
Attribute TableName.VB_Description = "Tabellenname über Tabellenindex"
    If m_cntTables = 0 Then Exit Property
    TableName = m_Tables(TableIndex).name
End Property
Public Property Get TableLine(ByVal TableIndex As Long, ByVal lRow As Long)
Attribute TableLine.VB_Description = "Zeiger auf Tabellenzeile"
    If m_cntTables = 0 Then Exit Property
    TableLine = ItGetLine(m_Tables(TableIndex).Handle, lRow)
End Property
Public Property Get Lines(ByVal TableIndex As Long) As Long
Attribute Lines.VB_Description = "Anzahl der Zeilen einer internen Tabelle"
    If m_cntTables = 0 Then Exit Property
    Lines = ItFill(m_Tables(TableIndex).Handle)
End Property
Public Property Get TableCount() As Long
Attribute TableCount.VB_Description = "Anzahl erfaßter Tabellen"
    TableCount = m_cntTables
End Property
Public Property Get ImportCount() As Long
Attribute ImportCount.VB_Description = "Anzahl erfaßter Importparameter"
    ImportCount = m_cntImport
End Property
Public Property Get ExportCount() As Long
Attribute ExportCount.VB_Description = "Anzahl erfaßter Exportparameter"
    ExportCount = m_cntExport
End Property
Public Property Get ExportParamValue(ByVal lIndex As Long) As String
Attribute ExportParamValue.VB_Description = "Wert des Exportparameters"
    If m_cntExport = 0 Then Exit Property
    ExportParamValue = m_Export(lIndex).Value
End Property
Public Property Get ExportValueByName(ByVal sParam As String) As String
Attribute ExportValueByName.VB_Description = "Wert des Exportparameters anhand des Namens ermitteln"
  Dim i As Long
    If m_cntExport = 0 Then Exit Property
    For i = 1 To UBound(m_Export())
        If m_Export(i).name = sParam Then
            ExportValueByName = m_Export(i).Value
            Exit For
        End If
    Next i
End Property
Public Property Get ImportParamValue(ByVal lIndex As Long) As String
Attribute ImportParamValue.VB_Description = "Wert eines Importparameters"
    If m_cntImport = 0 Then Exit Property
    ImportParamValue = m_Import(lIndex).Value
End Property
Public Property Get ImportValueByName(ByVal sParam As String) As String
Attribute ImportValueByName.VB_Description = "Wert eines Importparameters anhand des Namens ermitteln"
  Dim i As Long
    If m_cntImport = 0 Then Exit Property
    For i = 1 To UBound(m_Import())
        If m_Import(i).name = sParam Then
            ImportValueByName = m_Import(i).Value
            Exit For
        End If
    Next i
End Property
Public Property Get LineAsString(ByVal TableIndex As Long, ByVal Line As Long) As String
Attribute LineAsString.VB_Description = "Tabellenzeile als in Form einer Zeichenkette ausgeben"
  Dim lpTableLine As Long
  Dim Buffer As String
  Dim lLen As Long
    If m_cntTables = 0 Then Exit Property
    lpTableLine = ItGetLine(m_Tables(TableIndex).Handle, Line)
    lLen = ItLeng(m_Tables(TableIndex).Handle)
    Buffer = String$(lLen, vbNullChar)
    If lpTableLine <> 0 And lLen > 0 Then
        CopyMemoryA ByVal StrPtr(Buffer), ByVal lpTableLine, lLen
        LineAsString = Left$(StrConv(Buffer, vbUnicode), lLen / 2)
    End If
End Property

Public Sub Initialize()
Attribute Initialize.VB_Description = "Alle Rfc-Einstellungen initialisieren"
    CleanUp
End Sub
Public Sub CleanUp()
Attribute CleanUp.VB_Description = "Tabellen, Parameter und Speicherbereiche wieder freigeben"
  Dim i As Long

    If m_cntTables > 0 Then
        For i = 1 To UBound(m_Tables())
            If m_Tables(i).Handle > 0 Then
                m_Tables(i).Handle = ItFree(m_Tables(i).Handle)
                m_Tables(i).name = ""
            End If
        Next i
        Erase m_Tables()
    End If
    
    Erase m_Export()
    Erase m_Import()

    If m_hSpace <> 0 Then
        m_hSpace = RfcFreeParamSpace(m_hSpace)
    End If

    m_cntExport = 0
    m_cntImport = 0
    m_cntTables = 0

    bLocked = False

End Sub

Public Function AllocParamSpace(ByVal MaxEx As Long, ByVal MaxIm As Long, ByVal MaxTab As Long) As Long
Attribute AllocParamSpace.VB_Description = "Speicher für Parameterübergabe reservieren"
    If bLocked Then Exit Function
    m_hSpace = RfcAllocParamSpace(MaxEx, MaxIm, MaxTab)
    If m_hSpace > 0 Then
        bLocked = True
        If MaxTab > 0 Then
            ReDim m_Tables(1 To MaxTab) As TableProps
        End If
        If MaxEx > 0 Then
            ReDim m_Export(1 To MaxEx) As ParamProps
        End If
        If MaxIm > 0 Then
            ReDim m_Import(1 To MaxIm) As ParamProps
        End If
        m_MaxExport = MaxEx
        m_MaxImport = MaxIm
        m_MaxTables = MaxTab
    Else
        RaiseError EAllocateParamSpace
    End If
    AllocParamSpace = m_hSpace
End Function
Public Function AddExportParam(ByRef sParamName As String, ByRef sParamValue As String) As Long
Attribute AddExportParam.VB_Description = "Exportparameter festlegen"
  Dim Result As Long
    If Not bLocked Then Exit Function
    If m_cntExport = m_MaxExport Then
        RaiseError EAddExportParam, sParamName
        Exit Function
    End If
    Result = RfcAddExportParam(m_hSpace, m_cntExport, sParamName, Len(sParamName), 0, Len(sParamValue), ByVal sParamValue)
    If Result <> RFC_OK Then
        RaiseError EAddExportParam, sParamName
        Exit Function
    End If
    Incr m_cntExport
    m_Export(m_cntExport).name = sParamName
    m_Export(m_cntExport).Value = sParamValue
    AddExportParam = m_cntExport
End Function
Public Function AddImportParam(ByRef sParamName As String, ByRef sParamValue As String) As Long
Attribute AddImportParam.VB_Description = "Importparameter festlegen"
  Dim Result As Long
    If Not bLocked Then Exit Function
    If m_cntImport = m_MaxImport Then
        RaiseError EAddImportParam, sParamName
        Exit Function
    End If
    Result = RfcAddImportParam(m_hSpace, m_cntImport, sParamName, Len(sParamName), 0, Len(sParamValue), ByVal sParamValue)
    If Result <> RFC_OK Then
        RaiseError EAddImportParam, sParamName
        Exit Function
    End If
    Incr m_cntImport
    m_Import(m_cntImport).name = sParamName
    m_Import(m_cntImport).Value = sParamValue
    AddImportParam = m_cntImport
End Function
Public Function AddNewTable(ByRef sTabName As String, ByRef lLen As Long) As Long
Attribute AddNewTable.VB_Description = "Neue Tabelle anlegen und in den Übergabespeicher legen (nur eine Leertabelle)"
  Dim hIT As Long
  Dim Result As Long
    If Not bLocked Then Exit Function
    If m_cntTables = m_MaxTables Then
        RaiseError EAddNewTable, sTabName
        Exit Function
    End If
    hIT = ItCreate(sTabName, lLen, 0, 0)
    If hIT = 0 Then
        RaiseError EItCreate, sTabName
        Exit Function
    End If
    Result = RfcAddTable(m_hSpace, m_cntTables, sTabName, Len(sTabName), 0, lLen, hIT)
    If Result <> 0 Then
        RaiseError EAddNewTable, sTabName
        Exit Function
    End If
    Incr m_cntTables
    m_Tables(m_cntTables).Handle = hIT
    m_Tables(m_cntTables).name = sTabName
    AddNewTable = m_cntTables
End Function
Public Function CallReceiveExt(ByRef sFunction As String) As Long
Attribute CallReceiveExt.VB_Description = "Funktionsbaustein aufrufen"
  Dim Result As Long
  Dim ErrInfo As RFC_ERROR_INFO_EX
    If Not bLocked Then Exit Function
    Result = RfcCallReceiveExt(m_hRfc, m_hSpace, sFunction, m_Exception)
    If Result <> RFC_OK Then
        RfcLastErrorEx ErrInfo
        If Len(ErrInfo.Message) > 0 Then
            RaiseError ECallReceiveExt, Trim$(ErrInfo.Message) & " " & sFunction
        Else
            RaiseError ECallReceiveExt, m_Exception & " " & sFunction
        End If
        Exit Function
    End If
    CallReceiveExt = Result
End Function
Public Function CreateNewTable(ByVal sTabName As String, ByVal lLen As Long, Optional ByVal lOccurs As Long) As Long
Attribute CreateNewTable.VB_Description = "Neue Tabelle anlegen"
  Dim hIT As Long
    If Not bLocked Then Exit Function
    If m_cntTables = m_MaxTables Then
        RaiseError ECreateNewTable, sTabName
        Exit Function
    End If
    hIT = ItCreate(sTabName, lLen, lOccurs, 0)
    If hIT = 0 Then
        RaiseError ECreateNewTable, sTabName
        Exit Function
    End If
    Incr m_cntTables
    m_Tables(m_cntTables).Handle = hIT
    m_Tables(m_cntTables).name = sTabName
    CreateNewTable = m_cntTables
End Function
Public Function AddEntry(ByVal lTableIndex As Long, ByRef lpData As Long, ByRef lLen As Long) As Long
Attribute AddEntry.VB_Description = "Neuen Eintrag an interne Tabelle anhängen"
  Dim lpTableLine As Long
  Dim Buffer As String
    Buffer = Space$(lLen)
    CopyMemoryA ByVal StrPtr(Buffer), ByVal lpData, lLen
    lpTableLine = ItAppLine(m_Tables(lTableIndex).Handle)
    If lpTableLine = 0 Then
        RaiseError EAddEntry, m_Tables(lTableIndex).name
        Exit Function
    End If
    CopyMemoryA ByVal lpTableLine, ByVal StrPtr(StrConv(Buffer, vbFromUnicode)), lLen
    AddEntry = (lpTableLine <> 0)
End Function
Public Function AddTable(ByVal lTabIndex As Long, ByVal lLen As Long) As Long
Attribute AddTable.VB_Description = "Neue Tabelle anlegen (wird noch nicht in hSpace geschrieben!)"
  Dim Result As Long
    Result = RfcAddTable(m_hSpace, lTabIndex - 1, m_Tables(lTabIndex).name, Len(m_Tables(lTabIndex).name), 0, lLen, m_Tables(lTabIndex).Handle)
    AddTable = Result
    If Result <> RFC_OK Then
        RaiseError EAddTable, m_Tables(lTabIndex).name
        Exit Function
    End If
End Function

Private Sub RaiseError(ByRef EError As EErrorPosition, Optional ByRef sParam As String)
Attribute RaiseError.VB_Description = "Fehler auslösen"
    CleanUp
    RfcClose m_hRfc
    RaiseEvent Error(EError, sParam)
End Sub

Private Sub InsertLoginParam(ByVal Offset As Long, ByRef sParam As String)
Attribute InsertLoginParam.VB_Description = "Interne Funktion zur Erstellung des Connection-Strings"
  Dim lPos As Long
    lPos = StrPtr(m_ConString) + (Offset * 2)   '<-- BSTR
    CopyMemoryA ByVal lPos, ByVal StrPtr(sParam), LenB(sParam)
End Sub

Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Klasse initialisieren"
    m_ConString = "ASHOST=                                 SYSNR=   CLIENT=    USER=             PASSWD=         LANG=   TRACE=  ABAP_DEBUG= "
    bLocked = True
End Sub
Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "Klasse entladen"
    CleanUp
    If m_hRfc > 0 Then RfcClose m_hRfc
End Sub
