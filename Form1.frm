VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8760
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerSize 
      Interval        =   3000
      Left            =   7800
      Top             =   4950
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7770
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4762
            MinWidth        =   4762
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Picture         =   "Form1.frx":08CA
            Text            =   "Importando"
            TextSave        =   "Importando"
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   7770
      Top             =   4440
   End
   Begin MSComctlLib.ListView lvwCabeza 
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   1402
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   14111
      EndProperty
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6975
      Left            =   30
      TabIndex        =   2
      Top             =   810
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   529
      TabCaption(0)   =   "Usuarios"
      TabPicture(0)   =   "Form1.frx":0B6B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Option1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tablas"
      TabPicture(1)   =   "Form1.frx":0B87
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Option4"
      Tab(1).Control(1)=   "Option5"
      Tab(1).Control(2)=   "cmbUsuarios"
      Tab(1).Control(3)=   "ListViewTablas"
      Tab(1).Control(4)=   "lblUsuario"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Conexiones"
      TabPicture(2)   =   "Form1.frx":0BA3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListViewConexion"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tablespaces"
      TabPicture(3)   =   "Form1.frx":0BBF
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListViewTablespaces"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Ultimas SQL"
      TabPicture(4)   =   "Form1.frx":0BDB
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListViewSQL"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.OptionButton Option4 
         Height          =   195
         Left            =   -72210
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   810
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Height          =   195
         Left            =   -71190
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   810
         Width           =   255
      End
      Begin VB.ComboBox cmbUsuarios 
         Height          =   315
         Left            =   -74940
         TabIndex        =   7
         Top             =   390
         Width           =   3435
      End
      Begin VB.OptionButton Option3 
         Height          =   195
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   390
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Height          =   195
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   390
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   195
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   390
         Width           =   255
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   6600
         Left            =   30
         TabIndex        =   3
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Empresa"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Fecha Importación"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Version"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Sesión"
            Object.Width           =   1199
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "MB"
            Object.Width           =   1658
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewTablas 
         Height          =   6150
         Left            =   -74940
         TabIndex        =   8
         Top             =   750
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   10848
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TABLA"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "MB      "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "EXTENTS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "INITIAL_EXT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "NEXT_EXT"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "MAX_EXT"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewConexion 
         Height          =   6600
         Left            =   -74970
         TabIndex        =   12
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "OSuser"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Username"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Machine"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Program"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Serial"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewTablespaces 
         Height          =   6600
         Left            =   -74970
         TabIndex        =   13
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Tablespace"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "MB Tamaño"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "MB Usados"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "MB Libres"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fichero de datos"
            Object.Width           =   5468
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewSQL 
         Height          =   6600
         Left            =   -74970
         TabIndex        =   14
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Programa"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Fecha/Hora"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Usuario"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SQL"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label lblUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71310
         TabIndex        =   9
         Top             =   420
         Width           =   4935
      End
   End
   Begin VB.Menu MenuTabla 
      Caption         =   "MenuTabla"
      Begin VB.Menu mnuExpDMP 
         Caption         =   "Generar DMP"
         Index           =   0
      End
      Begin VB.Menu mnuImpDMP 
         Caption         =   "Importar DMP"
      End
      Begin VB.Menu mnuExp 
         Caption         =   "Generar Insert's"
         Index           =   1
      End
      Begin VB.Menu mnuImp 
         Caption         =   "Procesar Insert's"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SQL               As String
Private itmX              As ListItem
Private itmT              As ListItem
Private sngCoordenadaX    As Single
Private sngCoordenadaY    As Single
Private iorden            As Integer
Private iordenTablas      As Integer
Private rstGlobal         As ADODB.Recordset
Private cnnGlobal         As ADODB.Connection
Private fs                As Object
Private fNumber1          As Integer

Private Sub Form_Load()
         
10       On Error GoTo GestErr
         
20       Form1.Caption = "Monitor Oracle - Algoritmo S.A."
         
30       LlenarCabecera
         
40       Option1.Value = True
         
50       Exit Sub
GestErr:
60       Screen.MousePointer = vbNormal
70       MsgBox "[Form_Load]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub Inicio()
      Dim rst1    As ADODB.Recordset

10       On Error GoTo GestErr

20       StatusBar.Panels(1).Text = "Inicializando..."
30       Screen.MousePointer = vbHourglass
                  
40       Set cnnGlobal = New ADODB.Connection
50       cnnGlobal.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
60       cnnGlobal.Open
         
70       Set rstGlobal = New ADODB.Recordset
80       rstGlobal.CursorLocation = adUseClient
90       rstGlobal.LockType = adLockReadOnly
100      rstGlobal.CursorType = adOpenStatic
         
110      Set rst1 = New ADODB.Recordset
120      rst1.CursorLocation = adUseClient
130      rst1.LockType = adLockReadOnly
140      rst1.CursorType = adOpenStatic
         
150      SQL = " SELECT   DBA_USERS.USERNAME, DBA_USERS.CREATED, EMPRESAS.EMP_DESCRIPCION "
160      SQL = SQL & "    FROM DBA_USERS, "
170      SQL = SQL & "         EMPRESAS "
180      SQL = SQL & "   WHERE USERNAME LIKE 'SYSADMIN%' "
190      SQL = SQL & "     AND SUBSTR (DBA_USERS.USERNAME, 10, 3) = EMPRESAS.EMP_CODIGO_EMPRESA(+) "
200      Select Case iorden
            Case 1
210            SQL = SQL & "ORDER BY USERNAME "
220         Case 2
230            SQL = SQL & "ORDER BY EMP_DESCRIPCION "
240         Case 3
250            SQL = SQL & "ORDER BY CREATED DESC "
260         Case Else
270            SQL = SQL & "ORDER BY USERNAME "
280      End Select
         
290      rstGlobal.Open SQL, cnnGlobal
            
300      rstGlobal.MoveFirst
310      ListView.ListItems.Clear
320      Do While Not rstGlobal.EOF
330         Set itmX = ListView.ListItems.Add
            
340         itmX.SubItems(1) = IIf(IsNull(rstGlobal("USERNAME").Value), "", rstGlobal("USERNAME").Value)
350         itmX.SubItems(2) = IIf(IsNull(rstGlobal("EMP_DESCRIPCION").Value), "", rstGlobal("EMP_DESCRIPCION").Value)
360         itmX.SubItems(3) = IIf(IsNull(rstGlobal("CREATED").Value), "", rstGlobal("CREATED").Value)

            cmbUsuarios.AddItem IIf(IsNull(rstGlobal("USERNAME").Value), "", rstGlobal("USERNAME").Value)
            
370         SQL = "SELECT NRO_VERSION FROM " & rstGlobal("USERNAME").Value & ".VERSION_PRODUCTO"
380         Set rst1 = New ADODB.Recordset
390         rst1.CursorLocation = adUseClient
400         rst1.LockType = adLockReadOnly
410         rst1.CursorType = adOpenStatic
            
420         On Error Resume Next
430         rst1.Open SQL, cnnGlobal
440         If rst1.RecordCount > 0 Then
450            rst1.MoveFirst
460            itmX.SubItems(4) = IIf(IsNull(rst1("NRO_VERSION").Value), "", rst1("NRO_VERSION").Value)
470         End If
480         If Not rst1 Is Nothing Then
490            If rst1.State <> adStateClosed Then rst1.Close
500         End If
510         Set rst1 = Nothing
520         Err.Clear
530         On Error GoTo GestErr

540         rstGlobal.MoveNext
550      Loop
         
620      TimerRefresh_Timer
630      TimerSize.Enabled = True
         
640      Exit Sub
GestErr:
650      Screen.MousePointer = vbNormal
660      MsgBox "[Inicio]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub LlenarCabecera()
      Dim cnn     As ADODB.Connection
      Dim rst     As ADODB.Recordset

10       On Error GoTo GestErr
         
20       Set cnn = New ADODB.Connection
30       cnn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
40       cnn.Open
         
50       Set rst = New ADODB.Recordset
60       rst.CursorLocation = adUseClient
70       rst.LockType = adLockReadOnly
80       rst.CursorType = adOpenStatic
         
90       SQL = " SELECT BANNER "
100      SQL = SQL & "  FROM SYS.V_$VERSION "
110      SQL = SQL & " WHERE ROWNUM = 1 "
120      SQL = SQL & "UNION ALL "
130      SQL = SQL & "SELECT 'Instance_Name: ' || UPPER(INSTANCE_NAME) || '  Host_Name: ' || UPPER(HOST_NAME) || '  Version: ' || VERSION "
140      SQL = SQL & "       || '  Startup_Time: ' || TO_CHAR(STARTUP_TIME, 'DD/MM/YYYY') "
150      SQL = SQL & "  FROM V$INSTANCE "
160      SQL = SQL & "UNION ALL "
170      SQL = SQL & "SELECT 'Status: ' || STATUS || '  Shutdown_Pending: ' || SHUTDOWN_PENDING || '  Database_Status: ' || DATABASE_STATUS "
180      SQL = SQL & "  FROM V$INSTANCE "

190      rst.Open SQL, cnn
            
200      rst.MoveFirst
210      Do While Not rst.EOF
220         Set itmX = lvwCabeza.ListItems.Add
230         itmX.Text = rst("BANNER").Value
240         rst.MoveNext
250      Loop

260      cnn.Close
270      Set cnn = Nothing
280      If Not rst Is Nothing Then
290         If rst.State <> adStateClosed Then rst.Close
300      End If
310      Set rst = Nothing
         
320      Exit Sub
GestErr:
330      Screen.MousePointer = vbNormal
340      MsgBox "[LlenarCabecera]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Form_Terminate()
   End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub Option1_Click()
   iorden = 1
   Inicio
End Sub

Private Sub Option2_Click()
   iorden = 2
   Inicio
End Sub

Private Sub Option3_Click()
   iorden = 3
   Inicio
End Sub

Private Sub TimerRefresh_Timer()
         
10       On Error GoTo GestErr
         
20       StatusBar.Panels(1).Text = "Actualizando..."
30       Screen.MousePointer = vbHourglass
         
40       VerToad
         
50       VerImportacion
         
60       ListView.Refresh
         
70       TimerRefresh.Enabled = True
         
80       StatusBar.Panels(1).Text = ""
90       Screen.MousePointer = vbDefault
         
100      Exit Sub
GestErr:
110      Screen.MousePointer = vbNormal
120      MsgBox "[TimerRefresh_Timer]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub VerToad()
      Dim cnn     As ADODB.Connection
      Dim rst     As ADODB.Recordset
      Dim ix      As Integer

10       On Error GoTo GestErr
         
20       Set cnn = New ADODB.Connection
30       cnn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
40       cnn.Open
         
50       Set rst = New ADODB.Recordset
60       rst.CursorLocation = adUseClient
70       rst.LockType = adLockReadOnly
80       rst.CursorType = adOpenStatic
         
90       For Each itmX In ListView.ListItems
100         With itmX
110            SQL = " SELECT   COUNT (*) CONEXIONES, UPPER (V$SESSION.PROGRAM) PROGRAMA, V$SESSION.OSUSER, V$SESSION.USERNAME "
120            SQL = SQL & "    FROM V$SESSION "
130            SQL = SQL & "   WHERE USERNAME LIKE '" & .SubItems(1) & "' "
140            SQL = SQL & "     AND (   UPPER (V$SESSION.PROGRAM) = 'TOAD.EXE' "
150            SQL = SQL & "          OR UPPER (V$SESSION.PROGRAM) = 'DLLHOST.EXE') "
160            SQL = SQL & "GROUP BY SCHEMANAME, UPPER (PROGRAM), OSUSER, USERNAME "
170            SQL = SQL & "ORDER BY USERNAME, OSUSER "
180            rst.Open SQL, cnn
               
190            If rst.RecordCount > 0 Then
200               ix = 0
210               .ToolTipText = ""
220               Do While Not rst.EOF
230                  ix = ix + rst("CONEXIONES").Value
240                  .SubItems(5) = IIf(rst("PROGRAMA").Value = "TOAD.EXE", "Toad", "Dllhost")
250                  .ToolTipText = .ToolTipText & Right(rst("OSUSER").Value, Len(rst("OSUSER").Value) - InStr(rst("OSUSER").Value, ".")) & "/"
                     
260                  rst.MoveNext
270                  DoEvents
280               Loop
290               .ToolTipText = "(" & ix & ") " & .ToolTipText
300               .ToolTipText = Left(.ToolTipText, Len(.ToolTipText) - 1)
310            Else
320               .SubItems(5) = ""
330               .ToolTipText = ""
340            End If
350            rst.Close

360            DoEvents
370         End With
380      Next itmX
         
390      cnn.Close
400      Set cnn = Nothing
410      If Not rst Is Nothing Then
420         If rst.State <> adStateClosed Then rst.Close
430      End If
440      Set rst = Nothing
         
450      Exit Sub
GestErr:
460      Screen.MousePointer = vbNormal
470      MsgBox "[VerToad]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub VerImportacion()
      Dim cnn     As ADODB.Connection
      Dim rst     As ADODB.Recordset

10       On Error GoTo GestErr
         
20       Set cnn = New ADODB.Connection
30       cnn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
40       cnn.Open
         
50       Set rst = New ADODB.Recordset
60       rst.CursorLocation = adUseClient
70       rst.LockType = adLockReadOnly
80       rst.CursorType = adOpenStatic
         
90       SQL = " SELECT USERNAME, SCHEMANAME, OSUSER IMP_USUARIO "
100      SQL = SQL & "  FROM V$SESSION "
110      SQL = SQL & " WHERE USERNAME LIKE 'SYSADMIN_%' "
120      SQL = SQL & "   AND PROGRAM LIKE 'imp%' "
121      SQL = SQL & "   OR  PROGRAM LIKE 'IMP%' "
130      rst.Open SQL, cnn
         
131      If rst.RecordCount = 0 Then
            For Each itmX In ListView.ListItems
                With itmX
                   SetColor itmX, vbWindowText
                  DoEvents
               End With
            Next itmX

150         cnn.Close
160         Set cnn = Nothing
170         If Not rst Is Nothing Then
180            If rst.State <> adStateClosed Then rst.Close
190         End If
200         Set rst = Nothing
210         StatusBar.Panels(3).Text = "Ninguna Importación en Curso"
220         Form1.Caption = "Monitor Oracle - Algoritmo S.A."
230         Exit Sub
240      End If
            
250      For Each itmX In ListView.ListItems
260          With itmX
270            rst.Filter = "SCHEMANAME = '" & .SubItems(1) & "'"
280            If rst.RecordCount > 0 Then
290               .ToolTipText = rst("IMP_USUARIO").Value
300               SetColor itmX, vbRed
310               RecuperarTamaño rst("SCHEMANAME").Value
320            Else
330               SetColor itmX, vbWindowText
340            End If
350            rst.Filter = adFilterNone
360            DoEvents
370         End With
380      Next itmX
         
390      If rst.RecordCount = 1 Then
400         StatusBar.Panels(3).Text = rst.RecordCount & " Importación en Curso"
410      Else
420         StatusBar.Panels(3).Text = rst.RecordCount & " Importaciones en Curso"
430      End If
         
440      Form1.Caption = rst.RecordCount & " Imp. en Curso - Monitor Oracle - Algoritmo S.A."
         
450      cnn.Close
460      Set cnn = Nothing
470      If Not rst Is Nothing Then
480         If rst.State <> adStateClosed Then rst.Close
490      End If
500      Set rst = Nothing
         
510      Exit Sub
GestErr:
520      Screen.MousePointer = vbNormal
530      MsgBox "[VerImportacion]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub listview_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   sngCoordenadaX = x
   sngCoordenadaY = y
End Sub
Private Sub listviewTablas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   sngCoordenadaX = x
   sngCoordenadaY = y
End Sub
Private Sub ListViewConexion_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   sngCoordenadaX = x
   sngCoordenadaY = y
End Sub

Private Sub listview_Click()
   On Error Resume Next
   
   Set itmX = ListView.HitTest(sngCoordenadaX, sngCoordenadaY)
   If itmX Is Nothing Then Exit Sub
   StatusBar.Panels(2).Text = itmX.ToolTipText
   StatusBar.Panels(2).ToolTipText = itmX.ToolTipText
End Sub
Private Sub listview_KeyPress(KeyAscii As Integer)
   Set itmX = ListView.HitTest(sngCoordenadaX, sngCoordenadaY)
End Sub
Private Sub listviewTablas_KeyPress(KeyAscii As Integer)
   Set itmT = ListViewTablas.HitTest(sngCoordenadaX, sngCoordenadaY)
End Sub
Private Sub listview_KeyDown(KeyCode As Integer, Shift As Integer)
'   Set itmX = listview.HitTest(sngCoordenadaX, sngCoordenadaY)
'      If itmX Is Nothing Then Exit Sub
'      'ListView.SelectedItem.
'
'      'itmX.Selected
''   If listview.Items(itmX.Index).Focused = True Then
'      StatusBar.Panels(2).Text = itmX.ToolTipText
''   End If
End Sub
Private Sub listview_DblClick()
'      Dim strEjecuta As String

10       On Error GoTo GestErr

20       Set itmX = ListView.HitTest(sngCoordenadaX, sngCoordenadaY)
30       If itmX Is Nothing Then Exit Sub
         
         SSTab.Tab = 1
         cmbUsuarios.ListIndex = ComboSearch(cmbUsuarios, itmX.SubItems(1))
         
'40       strEjecuta = "C:\Archivos de programa\Quest Software\Toad for Oracle\toad.exe Connect=" & itmX.SubItems(1) & "/apfrms2001@CLIENTES"
'50       Shell strEjecuta, vbHide
         
60       Exit Sub
GestErr:
70       Screen.MousePointer = vbNormal
80       MsgBox "[listview_DblClick]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub SetColor(ByVal Item As MSComCtlLib.ListItem, ByVal Color As VBRUN.ColorConstants)
Dim l As ListSubItem
   Item.ForeColor = Color
   For Each l In Item.ListSubItems
      l.ForeColor = Color
   Next l
End Sub
Private Sub TimerSize_Timer()
   RecuperarTamaño
End Sub
Private Sub RecuperarTamaño(Optional ByVal strUsuario As String)
      Dim cnn     As ADODB.Connection
      Dim rst     As ADODB.Recordset

10       On Error GoTo GestErr
         
20       If Len(strUsuario) = 0 Then strUsuario = "SYSADMIN_%"
         
30       Set cnn = New ADODB.Connection
40       cnn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
50       cnn.Open
         
60       Set rst = New ADODB.Recordset
70       rst.CursorLocation = adUseClient
80       rst.LockType = adLockReadOnly
90       rst.CursorType = adOpenStatic
         
100      SQL = " SELECT   OWNER, TO_CHAR (ROUND (SUM (BYTES / 1024 / 1024), 2), '99,999.00') MB "
110      SQL = SQL & "    FROM DBA_SEGMENTS "
120      SQL = SQL & "   WHERE OWNER LIKE '" & strUsuario & "' "
130      SQL = SQL & " GROUP BY OWNER "
140      SQL = SQL & " ORDER BY OWNER "

150      rst.Open SQL, cnn
            
160      If rst.RecordCount > 0 Then
170         For Each itmX In ListView.ListItems
180             With itmX
190               rst.Filter = "OWNER = '" & .SubItems(1) & "'"
200               If rst.RecordCount > 0 Then
210                  itmX.SubItems(6) = IIf(IsNull(rst("MB").Value), "", rst("MB").Value)
220               End If
230               rst.Filter = adFilterNone
240               DoEvents
250            End With
260         Next itmX
270      End If

280      cnn.Close
290      Set cnn = Nothing
300      If Not rst Is Nothing Then
310         If rst.State <> adStateClosed Then rst.Close
320      End If
330      Set rst = Nothing
            
340      TimerSize.Enabled = False
350      ListView.Refresh
         
360      StatusBar.Panels(1).Text = ""
370      Screen.MousePointer = vbDefault
         
380      Exit Sub
GestErr:
390      Screen.MousePointer = vbNormal
400      MsgBox "[RecuperarTamaño]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub cmbUsuarios_Click()
      Dim cnn     As ADODB.Connection
      Dim rst     As ADODB.Recordset

10       On Error GoTo GestErr
         
20       If rstGlobal Is Nothing Then Exit Sub
         
30       rstGlobal.Filter = "USERNAME = '" & cmbUsuarios.Text & "'"
40       If rstGlobal.RecordCount > 0 Then
50          lblUsuario.Caption = IIf(IsNull(rstGlobal("EMP_DESCRIPCION").Value), "", rstGlobal("EMP_DESCRIPCION").Value)
60       End If

70       StatusBar.Panels(1).Text = "Actualizando Tablas..."
80       Screen.MousePointer = vbHourglass
                  
90       Set rst = New ADODB.Recordset
100      rst.CursorLocation = adUseClient
110      rst.LockType = adLockReadOnly
120      rst.CursorType = adOpenStatic
         
130      SQL = " SELECT  SEGMENT_NAME AS TABLA, (BYTES / 1024 / 1024) AS MB, EXTENTS, "
140      SQL = SQL & " (INITIAL_EXTENT / 1024) || 'K' AS INITIAL_EXT, (NEXT_EXTENT / 1024 / 1024) || 'MB' AS NEXT_EXT,"
150      SQL = SQL & " ROUND ((MAX_EXTENTS / 1024 / 1024), 2) || 'MB' AS MAX_EXT"
160      SQL = SQL & "    FROM DBA_SEGMENTS "
170      SQL = SQL & "   WHERE OWNER = '" & cmbUsuarios.Text & "'"
180      SQL = SQL & "     AND SEGMENT_TYPE = 'TABLE' "
190      Select Case iordenTablas
            Case 1
200            SQL = SQL & "ORDER BY TABLA "
210         Case 2
220            SQL = SQL & "ORDER BY BYTES DESC "
230         Case Else
240            SQL = SQL & "ORDER BY TABLA "
250      End Select
         
260      rst.Open SQL, cnnGlobal
            
270      rst.MoveFirst
280      ListViewTablas.ListItems.Clear
290      Do While Not rst.EOF
300         Set itmX = ListViewTablas.ListItems.Add
            
310         itmX.SubItems(1) = IIf(IsNull(rst("TABLA").Value), "", rst("TABLA").Value)
320         itmX.SubItems(2) = IIf(IsNull(rst("MB").Value), "", rst("MB").Value)
330         itmX.SubItems(3) = IIf(IsNull(rst("EXTENTS").Value), "", rst("EXTENTS").Value)
340         itmX.SubItems(4) = IIf(IsNull(rst("INITIAL_EXT").Value), "", rst("INITIAL_EXT").Value)
350         itmX.SubItems(5) = IIf(IsNull(rst("NEXT_EXT").Value), "", rst("NEXT_EXT").Value)
360         itmX.SubItems(6) = IIf(IsNull(rst("MAX_EXT").Value), "", rst("MAX_EXT").Value)

370         rst.MoveNext
380      Loop
         
390      If Not rst Is Nothing Then
400         If rst.State <> adStateClosed Then rst.Close
410      End If
420      Set rst = Nothing
430      Set rst = New ADODB.Recordset
         
440      rst.CursorLocation = adUseClient
450      rst.LockType = adLockReadOnly
460      rst.CursorType = adOpenStatic
         
470      StatusBar.Panels(1).Text = ""
         
480      SQL = " SELECT SUM (BYTES / 1024 / 1024) AS MB "
490      SQL = SQL & "  FROM DBA_SEGMENTS "
500      SQL = SQL & " WHERE OWNER = '" & cmbUsuarios.Text & "'"
510      SQL = SQL & "   AND SEGMENT_TYPE = 'TABLE' "
520      rst.Open SQL, cnnGlobal
530      If rst.RecordCount > 0 Then
540         StatusBar.Panels(2).Text = "Total: " & rst("MB").Value & " MB"
550      End If
         
560      Screen.MousePointer = vbDefault
570      rstGlobal.Filter = adFilterNone
         
580      Exit Sub
         
GestErr:
590      Screen.MousePointer = vbNormal
600      MsgBox "[cmbUsuarios_Click]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub Option4_Click()
   iordenTablas = 1
   cmbUsuarios_Click
End Sub

Private Sub Option5_Click()
   iordenTablas = 2
   cmbUsuarios_Click
End Sub

Private Sub ListViewTablas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

10       On Error GoTo GestErr

20       Set itmT = ListViewTablas.HitTest(sngCoordenadaX, sngCoordenadaY)
30       If itmT Is Nothing Then Exit Sub
         
40       If Button = 2 Then
50          PopupMenu MenuTabla
60       End If
         
70       Exit Sub
         
GestErr:
80       Screen.MousePointer = vbNormal
90       MsgBox "[ListViewTablas_MouseDown]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub mnuExpDMP_Click(Index As Integer)
         Dim strEjecuta As String
         Dim strNombre  As String
         
10       On Error GoTo GestErr
         
20       StatusBar.Panels(1).Text = "Generando DMP"
         
30       Set fs = CreateObject("Scripting.FileSystemObject")
         
40       If Not fs.folderexists("C:\MonitorOracle") Then
50          fs.CreateFolder "C:\MonitorOracle"
60       End If
         
70       strNombre = "C:\MonitorOracle\" & cmbUsuarios.Text & "_" & itmT.SubItems(1)
         
80       strEjecuta = "exp " & cmbUsuarios.Text & "/apfrms2001@CLIENTES file=" & strNombre & ".dmp log=" & strNombre & ".Exp.Log tables =(" & cmbUsuarios.Text & "." & itmT.SubItems(1) & ")"
90       Shell strEjecuta, vbHide
         
100      StatusBar.Panels(1).Text = ""
         
110      Exit Sub
         
GestErr:
120      Screen.MousePointer = vbNormal
130      StatusBar.Panels(1).Text = ""
140      MsgBox "[mnuExpDMP_Click]" & vbCrLf & Err.Description & Erl

End Sub
Private Sub mnuExp_Click(Index As Integer)
         Dim strEjecuta As String
         Dim strNombre  As String
         Dim strInserts As String
         Dim rst        As ADODB.Recordset
         
10       On Error GoTo GestErr
         
20       StatusBar.Panels(1).Text = "Generando Insert's"
         
30       Set fs = CreateObject("Scripting.FileSystemObject")
         
40       If Not fs.folderexists("C:\MonitorOracle") Then
50          fs.CreateFolder "C:\MonitorOracle"
60       End If
         
70       strNombre = "C:\MonitorOracle\" & cmbUsuarios.Text & "_" & itmT.SubItems(1) & ".sql"
80       strInserts = ""
         
90       Set cnnGlobal = New ADODB.Connection
100      cnnGlobal.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=" & cmbUsuarios.Text & ";Data Source=CLIENTES"
110      cnnGlobal.Open

120      Set rst = New ADODB.Recordset
130      rst.CursorLocation = adUseClient
140      rst.LockType = adLockReadOnly
150      rst.CursorType = adOpenStatic
         
160      SQL = " SELECT 'insert into ' || TABLE_NAME || ' (' "
170      SQL = SQL & "       || (SELECT RTRIM (EXTRACT (XMLAGG (XMLELEMENT (E, T.COLUMN_VALUE.GETROOTELEMENT () || ',')), "
180      SQL = SQL & "                                  '//text()'), "
190      SQL = SQL & "                         ',') "
200      SQL = SQL & "             FROM TABLE (XMLSEQUENCE (T.COLUMN_VALUE.EXTRACT ('ROW/*'))) T) "
210      SQL = SQL & "       || ') values (' "
220      SQL = SQL & "       || (SELECT DBMS_XMLGEN.CONVERT "
230      SQL = SQL & "                                  (RTRIM (EXTRACT (XMLAGG (XMLELEMENT (E, "
240      SQL = SQL & "                                                                       '''' "
250      SQL = SQL & "                                                                       || T.COLUMN_VALUE.EXTRACT ('//text()') "
260      SQL = SQL & "                                                                       || ''',')), "
270      SQL = SQL & "                                                   '//text()'), "
280      SQL = SQL & "                                          ','), "
290      SQL = SQL & "                                   1) "
300      SQL = SQL & "             FROM TABLE (XMLSEQUENCE (T.COLUMN_VALUE.EXTRACT ('ROW/*'))) T) "
310      SQL = SQL & "       || ');' INS_STMT "
320      SQL = SQL & "  FROM USER_TABLES, "
330      SQL = SQL & "       TABLE (XMLSEQUENCE (DBMS_XMLGEN.GETXMLTYPE ('select * from ' || TABLE_NAME).EXTRACT ('ROWSET/ROW'))) T "
340      SQL = SQL & " WHERE TABLE_NAME = '" & itmT.SubItems(1) & "' "
350      rst.Open SQL, cnnGlobal
         
360      If rst.RecordCount > 0 Then
370         rst.MoveFirst
380         Do While Not rst.EOF
390            strInserts = strInserts & rst("INS_STMT").Value
400            rst.MoveNext
               'If Not rst.EOF Then strInserts = strInserts & vbCrLf
410         Loop
            
420         fNumber1 = FreeFile
430         Open strNombre For Output As fNumber1
440         Print #fNumber1, Trim(strInserts)
450         Close #fNumber1
460      Else
470         MsgBox "Tabla sin Registros"
480      End If
         
490      StatusBar.Panels(1).Text = ""
         
500      Exit Sub
         
GestErr:
510      Screen.MousePointer = vbNormal
520      StatusBar.Panels(1).Text = ""
530      MsgBox "[mnuExp_Click]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub mnuImpDMP_Click()
   Dim strEjecuta As String
   Dim strNombre   As String

   On Error GoTo GestErr
   
'   strNombre = "C:\MonitorOracle\" & cmbUsuarios.Text & "_" & itmT.SubItems(1)
'   strEjecuta = "imp " & cmbUsuarios.Text & "/apfrms2001@CLIENTES FROMUSER=" & cmbUsuarios.Text & " TOUSER=" & cmbUsuarios.Text & " file=" & strNombre & ".dmp log=" & strNombre & ".Imp.Log tables =(" & itmT.SubItems(1) & ") INDEXES=N IGNORE=Y ROWS=y"
'   Shell strEjecuta, vbHide
   
   Exit Sub
   
GestErr:
   Screen.MousePointer = vbNormal
   MsgBox "[mnuImpDMP_Click]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub mnuImp_Click(Index As Integer)
         Dim strContents As String
         Dim strNombre   As String
         Dim objReadFile As Object
         Dim aInserts()  As String
         Dim ix          As Long

10       On Error GoTo GestErr
         
20       Set fs = CreateObject("Scripting.FileSystemObject")

         If cmbUsuarios.Text = "SYSADMIN" Then Exit Sub
         
30       strNombre = "C:\MonitorOracle\" & cmbUsuarios.Text & "_" & itmT.SubItems(1) & ".sql"
         
40       If fs.fileexists(strNombre) Then
50          Set objReadFile = fs.OpenTextFile(strNombre, 1)
60          strContents = objReadFile.ReadAll
70          objReadFile.Close
80          Set objReadFile = Nothing
90       End If
         
100      Set cnnGlobal = New ADODB.Connection
110      cnnGlobal.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=" & cmbUsuarios.Text & ";Data Source=CLIENTES"
120      cnnGlobal.Open
         
130      On Error Resume Next
140      aInserts = Split(Replace(strContents, vbCrLf, ""), ";")
150      For ix = LBound(aInserts) To UBound(aInserts)
160         cnnGlobal.Execute aInserts(ix)
170      Next ix
180      On Error GoTo GestErr
         
      '   strNombre = "C:\MonitorOracle\" & cmbUsuarios.Text & "_" & itmT.SubItems(1)
      '   strEjecuta = "imp " & cmbUsuarios.Text & "/apfrms2001@CLIENTES FROMUSER=" & cmbUsuarios.Text & " TOUSER=" & cmbUsuarios.Text & " file=" & strNombre & ".dmp log=" & strNombre & ".Imp.Log tables =(" & itmT.SubItems(1) & ") INDEXES=N IGNORE=Y ROWS=y"
      '   Shell strEjecuta, vbHide
         
190      Exit Sub
         
GestErr:
200      Screen.MousePointer = vbNormal
210      MsgBox "[mnuImp_Click]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub SSTab_Click(PreviousTab As Integer)
   Select Case SSTab.Tab
      Case Is = 0
      Case Is = 1
      Case Is = 2
         LlenarConexiones
      Case Is = 3
         LlenarTablespaces
      Case Is = 4
         LlenarSQL
   End Select
End Sub
Private Sub LlenarConexiones()
         Dim rst        As ADODB.Recordset
         
10       On Error GoTo GestErr
         
20       StatusBar.Panels(1).Text = "Actualizando Conexiones..."
30       Screen.MousePointer = vbHourglass
                  
40       Set rst = New ADODB.Recordset
50       rst.CursorLocation = adUseClient
60       rst.LockType = adLockReadOnly
70       rst.CursorType = adOpenStatic
         
80       SQL = " SELECT   OSUSER, USERNAME, MACHINE, PROGRAM, SID, SERIAL# "
90       SQL = SQL & "    FROM V$SESSION "
100      SQL = SQL & "   WHERE OSUSER <> 'oracle' "
110      SQL = SQL & "ORDER BY USERNAME "
120      rst.Open SQL, cnnGlobal
            
130      rst.MoveFirst
140      ListViewConexion.ListItems.Clear
150      Do While Not rst.EOF
160         Set itmX = ListViewConexion.ListItems.Add
            
170         itmX.SubItems(1) = IIf(IsNull(rst("osuser").Value), "", rst("osuser").Value)
180         itmX.SubItems(2) = IIf(IsNull(rst("username").Value), "", rst("username").Value)
190         itmX.SubItems(3) = IIf(IsNull(rst("machine").Value), "", Replace(rst("machine").Value, "ALGORITMO\", ""))
200         itmX.SubItems(4) = IIf(IsNull(rst("program").Value), "", rst("program").Value)
210         itmX.SubItems(5) = IIf(IsNull(rst("sid").Value), "", rst("sid").Value)
220         itmX.SubItems(6) = IIf(IsNull(rst("serial#").Value), "", rst("serial#").Value)

230         rst.MoveNext
240      Loop
         
250      StatusBar.Panels(1).Text = ""
260      Screen.MousePointer = vbNormal
         
270      Exit Sub
         
GestErr:
280      Screen.MousePointer = vbNormal
290      MsgBox "[LlenarConexiones]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub ListViewConexion_DblClick()
         
10       On Error GoTo GestErr

20       If MsgBox("¿Terminar Conexión?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
         
30       Set itmX = ListViewConexion.HitTest(sngCoordenadaX, sngCoordenadaY)
40       If itmX Is Nothing Then Exit Sub
         
50       SQL = "ALTER SYSTEM KILL SESSION '" & itmX.SubItems(5) & "," & itmX.SubItems(6) & "' IMMEDIATE"
60       cnnGlobal.Execute SQL
         
70       LlenarConexiones
         
80       Exit Sub
         
GestErr:
90       Screen.MousePointer = vbNormal
100      MsgBox "[ListViewConexion_DblClick]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub LlenarTablespaces()
         Dim rst        As ADODB.Recordset
         
10       On Error GoTo GestErr
         
20       If ListViewTablespaces.ListItems.Count > 0 Then Exit Sub
         
30       StatusBar.Panels(1).Text = "Actualizando Tablespaces..."
40       Screen.MousePointer = vbHourglass
                  
50       Set rst = New ADODB.Recordset
60       rst.CursorLocation = adUseClient
70       rst.LockType = adLockReadOnly
80       rst.CursorType = adOpenStatic
         
90       SQL = " SELECT   T.TABLESPACE_NAME ""Tablespace"", ROUND (MAX (D.BYTES) / 1024 / 1024, 2) ""MB Tamaño"", "
100      SQL = SQL & "         ROUND ((MAX (D.BYTES) / 1024 / 1024) - (SUM (DECODE (F.BYTES, NULL, 0, F.BYTES)) / 1024 / 1024), "
110      SQL = SQL & "                2) ""MB Usados"", ROUND (SUM (DECODE (F.BYTES, NULL, 0, F.BYTES)) / 1024 / 1024, 2) ""MB Libres"", "
120      SQL = SQL & "         SUBSTR (D.FILE_NAME, 1, 80) ""Fichero de datos"" "
130      SQL = SQL & "    FROM DBA_FREE_SPACE F, DBA_DATA_FILES D, DBA_TABLESPACES T "
140      SQL = SQL & "   WHERE T.TABLESPACE_NAME = D.TABLESPACE_NAME "
150      SQL = SQL & "     AND F.TABLESPACE_NAME(+) = D.TABLESPACE_NAME "
160      SQL = SQL & "     AND F.FILE_ID(+) = D.FILE_ID "
170      SQL = SQL & "GROUP BY T.TABLESPACE_NAME, D.FILE_NAME, T.PCT_INCREASE, T.STATUS "
180      SQL = SQL & "ORDER BY 1, 4 "
190      rst.Open SQL, cnnGlobal
            
200      rst.MoveFirst
210      ListViewTablespaces.ListItems.Clear
220      Do While Not rst.EOF
230         Set itmX = ListViewTablespaces.ListItems.Add
            
240         itmX.SubItems(1) = IIf(IsNull(rst("Tablespace").Value), "", rst("Tablespace").Value)
250         itmX.SubItems(2) = IIf(IsNull(rst("MB Tamaño").Value), "", rst("MB Tamaño").Value)
260         itmX.SubItems(3) = IIf(IsNull(rst("MB Usados").Value), "", rst("MB Usados").Value)
270         itmX.SubItems(4) = IIf(IsNull(rst("MB Libres").Value), "", rst("MB Libres").Value)
280         itmX.SubItems(5) = IIf(IsNull(rst("Fichero de datos").Value), "", rst("Fichero de datos").Value)

290         rst.MoveNext
300      Loop
         
310      StatusBar.Panels(1).Text = ""
320      Screen.MousePointer = vbNormal
         
330      Exit Sub
         
GestErr:
340      Screen.MousePointer = vbNormal
350      MsgBox "[LlenarTablespaces]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub ListViewSQL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Set itmX = ListViewSQL.HitTest(x, y)
   If itmX Is Nothing Then Exit Sub
   ListViewSQL.ToolTipText = itmX.ToolTipText
End Sub
Private Sub LlenarSQL()
         Dim rst        As ADODB.Recordset
         
10       On Error GoTo GestErr
         
20       StatusBar.Panels(1).Text = "Actualizando SQL's..."
30       Screen.MousePointer = vbHourglass
                  
40       Set rst = New ADODB.Recordset
50       rst.CursorLocation = adUseClient
60       rst.LockType = adLockReadOnly
70       rst.CursorType = adOpenStatic
         
80       SQL = " SELECT DISTINCT VS.MODULE, "
90       SQL = SQL & "                TO_CHAR (TO_DATE (VS.FIRST_LOAD_TIME, 'YYYY-MM-DD/HH24:MI:SS'), 'MM/DD  HH24:MI:SS') FIRST_LOAD_TIME, "
100      SQL = SQL & "                AU.USERNAME PARSEUSER, VS.SQL_TEXT "
110      SQL = SQL & "           FROM V$SQLAREA VS, ALL_USERS AU "
120      SQL = SQL & "          WHERE (PARSING_USER_ID != 0) "
130      SQL = SQL & "            AND (AU.USER_ID(+) = VS.PARSING_USER_ID) "
140      SQL = SQL & "            AND (EXECUTIONS >= 1) "
150      SQL = SQL & "            AND ROWNUM <= 50 "
160      SQL = SQL & "       ORDER BY 2 DESC "
170      rst.Open SQL, cnnGlobal
            
180      rst.MoveFirst
190      ListViewSQL.ListItems.Clear
200      Do While Not rst.EOF
210         Set itmX = ListViewSQL.ListItems.Add
            
220         itmX.SubItems(1) = IIf(IsNull(rst("MODULE").Value), "", rst("MODULE").Value)
230         itmX.SubItems(2) = IIf(IsNull(rst("FIRST_LOAD_TIME").Value), "", rst("FIRST_LOAD_TIME").Value)
240         itmX.SubItems(3) = IIf(IsNull(rst("PARSEUSER").Value), "", rst("PARSEUSER").Value)
250         itmX.SubItems(4) = IIf(IsNull(rst("SQL_TEXT").Value), "", rst("SQL_TEXT").Value)
260         itmX.ToolTipText = IIf(IsNull(rst("SQL_TEXT").Value), "", rst("SQL_TEXT").Value)

270         rst.MoveNext
280      Loop
         
290      StatusBar.Panels(1).Text = ""
300      Screen.MousePointer = vbNormal
         
310      Exit Sub
         
GestErr:
320      Screen.MousePointer = vbNormal
330      MsgBox "[LlenarSQL]" & vbCrLf & Err.Description & Erl
End Sub



'Mostrar un dmp
'imp SYSADMIN/apfrms2001@CLIENTES SHOW=Y FULL=Y file=D:\AIBAL.DMP
