'   version: 2021-09-09
'   Created by Miroslav Suba
'   msuba@hpe.com

Option Explicit
Option Private Module

Public Const error_SAP_Disconnected = -2147417848
Public Const error_SAP_Logon_EntryNotFound = 1000

Public Const fsForReading = 1
Public Const fsForWriting = 2
Public Const fsForAppending = 8

Private Const moduleVersion = "Q3JlYXRlZCBieSBtc3ViYUBocGUuY29t"

Private Type T_SAP_Client
    systemID As String
    userName As String
End Type

Public SAPRot As Object
Public SAPGUIAuto As Object
Public SAPApp As Object
Public SAPConnection As Object

Public filePathSaveAs As String
Public sessionWasLoggedByMacro As Boolean
Public exportTimeOut As Long

'SAP GUI Tree structure for IDocs
'---------------------------------------------------
'   Node Key                               Node Path
'---------------------------------------------------
'   IDoc Selection                         1
'     Idoc Number is equal                  1\1
'   BCSO Development System D01            2
'     Idoc in inbound processing            2\1
'      Application document not posted      2\1\1

Public Type T_SAP_TreeNode
    nodeKey As String
    nodePath As String
    nodeItems As Variant
End Type

Public Type T_SAP_TreeColumn
    columnName As String
    columnTitle As String
End Type

Public Type T_SAP_Tree
    SID As String

    columns() As T_SAP_TreeColumn
    listTreeNodes() As T_SAP_TreeNode
End Type

'---
'   Speed optimization for functions:
'   SAP_LoadAllObjects - used by SAP_GetValidObjectID
'
'   passportTransactionID helps us to identify if transaction was changed
'   T_SAP_ObjectID - properties of object ID
'   listAllSID - all SAP object IDs loaded by SAP_LoadAllObjects
'---

Public passportTransactionID As String

Public Type T_SAP_ObjectID
    ID As String
    
    textValue As String
    typeValue As String
    nameValue As String
    
    changeAble As Boolean
    containerType As Boolean
End Type

Public listAllSID() As T_SAP_ObjectID

Public SAPSystemID As Long

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   SOME INTERNAL FUNCTIONS
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub FSO_DeleteFile(fileName As String)
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Force file deletion, in case file exists already
    If fso.FileExists(fileName) Then
        'object.DeleteFile filespec, [ force ]
        fso.DeleteFile fileName, True
    End If
End Sub

Private Function formatAsFolderPath(ByVal folderPath As String) As String
    folderPath = Trim(folderPath)
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    formatAsFolderPath = folderPath
End Function

Private Function getFileExtension(ByVal fileName As String) As String
    Dim I As Long
    
    getFileExtension = ""
    
    If InStr(fileName, ".") > 0 Then
        For I = Len(fileName) To 1 Step -1
            If Mid(fileName, I, 1) = "." Then
                getFileExtension = Mid(fileName, I + 1, Len(fileName))
                Exit Function
            End If
        Next I
    End If
End Function

Function changeExtension(ByVal fileName As String, fileExtension As String)
    Dim I As Long
    
    If InStr(fileName, ".") > 0 Then
        For I = Len(fileName) To 1 Step -1
            If Mid(fileName, I, 1) = "." Then
                fileName = Mid(fileName, 1, I - 1)
                Exit For
            End If
        Next I
    End If

    changeExtension = fileName & "." & fileExtension
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will check if string s is in array v
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function stringIsInArray(ByVal s As String, ByVal v As Variant, Optional caseSensitive As Boolean = False) As Boolean
    Dim I As Long
    
    stringIsInArray = False
    
    If IsArray(v) Then
        For I = LBound(v) To UBound(v)
            If caseSensitive Then
                If s = v(I) Then
                    stringIsInArray = True
                    Exit Function
                End If
            Else
                If UCase(s) = UCase(v(I)) Then
                    stringIsInArray = True
                    Exit Function
                End If
            End If
        Next I
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will check if object o is nothing, if yes - it will display errorMsg
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function invalidObject(ByVal o As Object, ByVal errorMsg As String) As Boolean
    invalidObject = False
    
    If o Is Nothing Then
        invalidObject = True
        MsgBox errorMsg, vbCritical, "SAP Initialization Error"
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will check if v is numeric value, returns vbNullString if not, otherwise v converted to number
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function getNumericValue(ByVal v As Variant) As Variant
    'Convert to string, remove extra spaces
    v = CStr(v)
    v = Trim(v)
    
    If IsNumeric(v) Then
        v = CLng(v)
    Else
        v = vbNullString
    End If
    
    getNumericValue = v
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Safe function to get changeAble property (not all objects have this one!)
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function SAP_GetChangeAble(o As Object) As Boolean
    Dim listGUIwithChangeAbleProperty As Variant

    SAP_GetChangeAble = False
    
    If o Is Nothing Then Exit Function
    
'    Dim listGUIComponentType As Variant
'    listGUIComponentType = Array("GuiAbapEditor", "GuiApoGrid", "GuiApplication", "GuiBarChart", "GuiBox", "GuiButton", "GuiCalendar", "GuiChart", "GuiCheckBox", "GuiCollection", "GuiColorSelector", "GuiComboBox", "GuiComboBoxControl", "GuiComboBoxEntry", "GuiComponent", _
'                                  "GuiComponentCollection", "GuiConnection", "GuiContainer", "GuiContainerShell", "GuiContextMenu", "GuiCTextField", "GuiCustomControl", "GuiDialogShell", "GuiEAIViewer2D", "GuiEAIViewer3D", "GuiEnum", "GuiFrameWindow", "GuiGOSShell", "GuiGraphAdapt", _
'                                  "GuiGridView", "GuiHTMLViewer", "GuiInputFieldControl", "GuiLabel", "GuiMainWindow", "GuiMap", "GuiMenu", "GuiMenubar", "GuiMessageWindow", "GuiModalWindow", "GuiNetChart", "GuiOfficeIntegration", "GuiOkCodeField", "GuiPasswordField", "GuiPicture", "GuiRadioButton", _
'                                  "GuiSapChart", "GuiScrollbar", "GuiScrollContainer", "GuiSession", "GuiSessionInfo", "GuiShell", "GuiSimpleContainer", "GuiSplit", "GuiSplitterContainer", "GuiStage", "GuiStatusbar", "GuiStatusPane", "GuiTab", "GuiTableColumn", "GuiTableControl", _
'                                  "GuiTableRow", "GuiTabStrip", "GuiTextedit", "GuiTextField", "GuiTitlebar", "GuiToolbar", "GuiToolbarControl", "GuiTree", "GuiUserArea", "GuiUtils", "GuiVComponent", "GuiVContainer")
    
    'TODO: sort list by usage of objects (will speed up process, I already sorted it a little bit, but there is more to do :))
    
    'Object doesn't support this property or method
    'GuiTitlebar
    
    listGUIwithChangeAbleProperty = Array("GuiMenu", "GuiButton", "GuiLabel", "GuiTextField", "GuiCTextField", "GuiStatusPane", "GuiTab", "GuiComboBox", "GuiToolbar", "GuiShell", "GuiCheckBox", "GuiContainerShell", "GuiSimpleContainer", _
                                          "GuiMenubar", "GuiStatusbar", "GuiUserArea", "GuiOkCodeField", "GuiBox", "GuiScrollContainer", "GuiGOSShell", "GuiTableControl", "GuiCustomControl", "GuiDialogShell", _
                                          "GuiEAIViewer2D", "GuiEAIViewer3D", "GuiFrameWindow", "GuiGraphAdapt", "GuiGridView", "GuiHTMLViewer", "GuiMainWindow", "GuiMap", "GuiMessageWindow", _
                                          "GuiModalWindow", "GuiNetChart", "GuiOfficeIntegration", "GuiPasswordField", "GuiRadioButton", "GuiSapChart", "GuiSplit", "GuiSplitterContainer", _
                                          "GuiStage", "GuiTextedit", "GuiToolbarControl", "GuiTree", "GuiVComponent", "GuiVContainer", "GuiAbapEditor", "GuiApoGrid", "GuiBarChart", "GuiCalendar", "GuiChart", "GuiColorSelector", "GuiContextMenu")
    
    If stringIsInArray(Trim(o.Type), listGUIwithChangeAbleProperty) Then
        'MsgBox o.type
        SAP_GetChangeAble = o.changeAble
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function converts nodePath
'       1\1\1 --> 1\2
'       1\2\1 --> 1\3
'       and so on ...
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getNextNodePathParent(ByVal nodePath As String) As String
    Dim I As Long

    Dim nodeList() As String
    Dim nodePathNew As String
    
    nodePathNew = ""
    
    If InStr(nodePath, "\") > 0 Then
        nodeList = Split(nodePath, "\")
        
        If UBound(nodeList) > 1 Then
            For I = LBound(nodeList) To UBound(nodeList) - 2
                If nodePathNew > "" Then nodePathNew = nodePathNew & "\"
                nodePathNew = nodePathNew & nodeList(I)
            Next I
            
            getNextNodePathParent = nodePathNew & "\" & Val(nodeList(UBound(nodeList) - 1)) + 1
        Else
            getNextNodePathParent = Val(nodeList(UBound(nodeList) - 1)) + 1
        End If
    Else
        getNextNodePathParent = Val(nodePath) + 1
    End If
    
    Erase nodeList
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function converts nodePath
'       1\1\1 --> 1\1\2
'       1\2\1 --> 1\2\2
'       and so on ...
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getNextNodePath(ByVal nodePath As String) As String
    Dim I As Long
    Dim nodeList() As String
    Dim nodePathNew As String
    
    nodePathNew = ""
    
    If InStr(nodePath, "\") > 0 Then
        nodeList = Split(nodePath, "\")
        
        For I = 0 To UBound(nodeList) - 1
            If nodePathNew > "" Then nodePathNew = nodePathNew & "\"
            
            nodePathNew = nodePathNew & nodeList(I)
        Next I
        
        getNextNodePath = nodePathNew & "\" & Val(nodeList(UBound(nodeList))) + 1
    Else
        getNextNodePath = Val(nodePath) + 1
    End If
    
    Erase nodeList
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   PUBLIC FUNCTIONS
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will extract 'clean' SAP object ID
'   /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VBELN --> wnd[0]/usr/ctxtVBAK-VBELN
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_GetSID(ByVal SID As String, Optional ByVal startsFrom = "wnd[") As String
    Dim I As Long
    
    I = InStr(SID, startsFrom)
    If I > 0 Then SID = Mid(SID, I)
    
    SAP_GetSID = SID
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will extrat last part of SAP object ID
'   /app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VBELN --> ctxtVBAK-VBELN
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_GetLastSID(ByVal SID As String) As String
    Do
        If InStr(SID, "/") > 0 Then
            SID = Mid(SID, InStr(SID, "/") + 1)
        End If
    Loop While InStr(SID, "/") > 0
    
    SAP_GetLastSID = SID
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will get row number from SAP object ID - SID
'   wnd[0]/usr/lbl[91,15] --> 15
'                 [column, row]
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_GetSIDRow(ByVal SID As String) As Long
    Dim I As Long
    
    SAP_GetSIDRow = -1
    
    SID = SAP_GetLastSID(SID)
    
    I = InStr(SID, "lbl[")
    If I > 0 Then
        SID = Mid(SID, I + 4)
    
        I = InStr(SID, ",")
        If I > 0 Then
            SID = Mid(SID, I + 1)
    
            I = InStr(SID, "]")
            If I > 0 Then
                SID = Mid(SID, 1, I - 1)
                
                If IsNumeric(SID) Then SAP_GetSIDRow = Val(SID)
            End If
        End If
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will get column number from SAP object ID - SID
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_GetSIDCol(ByVal SID As String) As Long
    Dim I As Long
    
    SAP_GetSIDCol = -1
    
    SID = SAP_GetLastSID(SID)

    I = InStr(SID, "lbl[")
    If I > 0 Then
        SID = Mid(SID, I + 4)
    
        I = InStr(SID, ",")
        If I > 0 Then
            SID = Mid(SID, 1, I - 1)
    
            If IsNumeric(SID) Then SAP_GetSIDCol = Val(SID)
        End If
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Releases all objects from memory used by this module
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub SAP_Destroy(vSession As Object)
    Set vSession = Nothing
    
    'Global SAP variables
    Set SAPRot = Nothing
    Set SAPGUIAuto = Nothing
    Set SAPApp = Nothing
    Set SAPConnection = Nothing
    
    passportTransactionID = ""
    
    Erase listAllSID
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Sets filePathSaveAs to Temporary files if not specified by filePath
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub SAP_SetFilePathSaveAs(Optional ByVal filePath As String = "")
    If filePath = "" Then
        filePathSaveAs = Environ("TEMP")
    Else
        filePathSaveAs = filePath
    End If
    
    filePathSaveAs = formatAsFolderPath(filePathSaveAs)
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Will get text from Status bar
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_GetStatusBar(vSession As Object) As String
    Dim s As String
    
    s = vSession.FindByID("wnd[0]/sbar").Text
    s = Trim(s)

    SAP_GetStatusBar = s
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Will get text from Title
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_GetTitleBar(vSession As Object) As String
    Dim s As String
    
    s = vSession.FindByID("wnd[0]/titl").Text
    s = Trim(s)

    SAP_GetTitleBar = s
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Will get window ID - winNo = 1 (first 'sub window', e.g. Find Variant)
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_GetWindowID(vSession As Object, Optional winNo As Long = 0) As String
    SAP_GetWindowID = ""
    
    If vSession.Children.Count > winNo Then
        'winNo has to be converted to Integer
        SAP_GetWindowID = vSession.Children(CInt(winNo)).ID
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Will clear all text fields in Session, in specified area, by default -> wnd[0]/usr
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub SAP_ClearAllTextFields(vSession As Object, Optional ByVal searchArea = "wnd[0]/usr")
    Dim o As Object
    
    'Clear all text fields
    For Each o In vSession.FindByID(searchArea).Children
        'Text fields only
        If stringIsInArray(Trim(o.Type), Array("GuiTextField", "GuiCTextField")) Then
            'If ChangeAble
            If o.changeAble = True Then o.Text = ""
        End If
    Next o
    
    Set o = Nothing
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Will get the number of Session, where tCode is active, if there is no such session -1 value will be returned
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_GetSessionNoByTCode(vSession As Object, tCode As String) As Long
    Dim I As Long
    
    SAP_GetSessionNoByTCode = -1
    
    tCode = UCase(Trim(tCode))
    
    'Empty Session
    Select Case tCode
    Case "0", "SAP EASY ACCESS":
        tCode = "SESSION_MANAGER"
    End Select
    
    'Init SAP if not done already
    If SAPConnection Is Nothing Then SAP_Init vSession
    
    If invalidObject(SAPConnection, "Error while initializing object SAPApp.Children() of GetScriptingEngine") = False Then
        For I = 1 To SAPConnection.Children.Count
            'SessionNo has to be converted to Integer
            Set vSession = SAPConnection.Children(CInt(I - 1))
            
            If UCase(Trim(vSession.Info.Transaction)) = tCode Then
                SAP_GetSessionNoByTCode = I
                Exit Function
            End If
        Next I
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Will log in to SAP session via specified connectionName, with specified userName and password
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub SAP_OpenConnection(vSession As Object, ByVal connectionName As String, Optional userName As String = "", Optional password As String = "", Optional synchronousMode As Boolean = True)
    Dim flagSAPLogonEntryNotFound As Boolean
    
    flagSAPLogonEntryNotFound = False
    
    Set SAPGUIAuto = GetObject("SAPGUI")
    'Set SAPGUIAuto = CreateObject("SAPGUI")
    'Set SAPGUIAuto = CreateObject("SAPGUI.ScriptingCtrl.1")

    Set SAPApp = SAPGUIAuto.GetScriptingEngine()
    
    On Error Resume Next
    
    Set SAPConnection = SAPApp.OpenConnection(connectionName, synchronousMode)
    
    If Err.Number = error_SAP_Logon_EntryNotFound Then
        flagSAPLogonEntryNotFound = True
        MsgBox "SAP Logon connection entry not found " & connectionName, vbExclamation
    End If

    If Err.Number <> 0 Then Err.Clear
    
    On Error GoTo -1
    
    If flagSAPLogonEntryNotFound Then Exit Sub

'---
    
    Set vSession = SAPConnection.Children(0)

    'Wait for window to appear
    Do
        Application.Wait (Now + TimeValue("00:00:01"))
        DoEvents
    Loop While vSession.Info.Transaction <> "S000"
    
    'If username and password are specified - enter them
    If userName <> "" Then vSession.FindByID("wnd[0]/usr/txtRSYST-BNAME").Text = userName
    If password <> "" Then vSession.FindByID("wnd[0]/usr/pwdRSYST-BCODE").Text = password

    If userName <> "" And password <> "" Then vSession.FindByID("wnd[0]/tbar[0]/btn[0]").Press
    
    Do
        Application.Wait (Now + TimeValue("00:00:01"))
        
        'License Information for Multiple Logon
        Dim SID As String
        SID = SAP_GetWindowID(vSession, 1)
        
        If SID <> "" Then
            If vSession.FindByID(SID).Text = "License Information for Multiple Logon" Then
                'TODO: macro waits for user action
            End If
        End If
        
        DoEvents
    Loop While UCase(Trim(vSession.Info.Transaction)) <> "SESSION_MANAGER"
    
    sessionWasLoggedByMacro = True
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Method will connect vSession object to SAP session
'   by specifying systemID you can connect to specific system (eg: R01, DCP). If not specified user will be prompted to select system manually
'   by specifying sessionNo you can connect to specific session
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub SAP_Init(vSession As Object, Optional systemID As String = "", Optional sessionNo As Long = -1)
    Dim I As Long
    
    Dim SID As String
    
    Dim listClients() As T_SAP_Client
    Dim listClientsCount As Long
    
    Dim sessionsCount As Long
    
    Dim title As String
    Dim msg As String
    Dim response As Variant

    Dim defaultClient As String

    SAP_Destroy vSession
    
    'export time out 1 hour in case of MHTML files
    exportTimeOut = 3600
    
    On Error Resume Next
    
    'Try to connect using SAP ROT Wrapper
    Set SAPRot = CreateObject("SapROTWr.SapROTWrapper")
    Set SAPGUIAuto = SAPRot.GetROTEntry("SAPGUI")
    
    Err.Clear
    
    'If it does not work this way, try default
    If SAPGUIAuto Is Nothing Then
        Set SAPGUIAuto = GetObject("SAPGUI")
        
        If invalidObject(SAPGUIAuto, "Error while initializing object SAPGUI." & Chr(10) & "Please make sure SAP Logon Launchpad is running.") Then GoTo Error_Handler
    End If
    
    Set SAPApp = SAPGUIAuto.GetScriptingEngine()
    If invalidObject(SAPApp, "Error while initializing object GetScriptingEngine of SAPGUI.") Then GoTo Error_Handler
    
'Detect Clients
    
    listClientsCount = SAPApp.Children.Count
    
    If listClientsCount = 0 Then
        title = "SAP Initialization Error"
        
        If systemID <> "" Then
            title = title & " (" & systemID & ")"
        End If

        response = MsgBox("You are not logged in." & Chr(10) & "Do you want to log in now?", vbCritical + vbYesNo, title)
        If response = vbYes Then GoTo Open_New_Connection
        GoTo Exit_Program
    End If
    
    ReDim listClients(listClientsCount)
    
    For I = 1 To listClientsCount
        'We will ignore busy clients
        If SAPApp.Children(CInt(I - 1)).Children(0).Busy = False Then
            'SessionNo has to be converted to Integer
            listClients(I).systemID = Trim(SAPApp.Children(CInt(I - 1)).Children(0).Info.SystemName)
            listClients(I).userName = Trim(SAPApp.Children(CInt(I - 1)).Children(0).Info.user)
        Else
            listClients(I).systemID = "[session is busy]"
        End If
    Next I
    
    'if scripting is not enabled - Info.User will not be available and will raise an error
    Err.Clear
    
    msg = ""
    response = ""
    defaultClient = "1"
    
    'If user specified to which Client he wants to connect with
    If systemID <> "" Then
        For I = 0 To listClientsCount
            If listClients(I).systemID Like systemID Then
                response = I
                Exit For
            End If
        Next I
        
        If response = "" Then
            msg = "Client " & systemID & " not detected." & Chr(10)
            defaultClient = "+"
        End If
    End If
    
    'In case that more then one client is available
    If response = "" Then
        If listClientsCount > 1 Then
            msg = msg & "More than one SAP connection is available." & Chr(10)
        End If
    End If
    
    'If there is any 'warning' msg in variable msg then User has to select with which client he would like to be connected
    '(either he did not specify client and we have more of them available, or he specified a different client than the one which is available)
    If msg <> "" Then
        msg = msg & "Select client, you would like to connect with:"
        
        'Create a list of clients
        For I = 1 To listClientsCount
            msg = msg & Chr(10) & I & " - " & listClients(I).systemID & " " & listClients(I).userName
        Next I
        
        msg = msg & Chr(10) & "+ - to open new connection."
            
        response = InputBox(msg, "SAP Initialization", defaultClient)
        
        '+ open new connection
        If response = "+" Then
        
Open_New_Connection:
        
            response = InputBox("Please enter Logon entry name:", "SAP Logon connection entry", systemID)
            
            If response <> "" Then
               'User will have to enter
                SAP_OpenConnection vSession, response
                
                If SAP_Activated(vSession, systemID) = False Then GoTo Exit_Program
                
                Set SAPConnection = vSession.Parent
                
                GoTo New_Connection_Opened
            End If
        End If
        
        response = getNumericValue(response)
                
        If response = "" Then GoTo Exit_Program
        
        'Wrong input ?
        If (response < 1) Or (response > listClientsCount) Then
            response = ""
            Call MsgBox("Wrong input !", vbCritical, "SAP Initialization Error")
        End If
    Else
        'And of course, if there is one client available, we will connect to that one
        If response = "" Then response = 1
    End If
        
    Set SAPConnection = SAPApp.Children(response - 1)
    
    If invalidObject(SAPConnection, "Error while initializing object SAPApp.Children() of GetScriptingEngine") Then GoTo Error_Handler
    
    If SAPConnection.DisabledByServer = True Then
        Call MsgBox("Scripting support has not been enabled for the application server." & Chr(10) & _
                    Trim(SAPConnection.connectionstring), vbCritical, "SAP Initialization error")
                    
        GoTo Exit_Program
    End If
    
    On Error GoTo -1
    
'Detect Session no (we can create new session with this function)

New_Connection_Opened:

    sessionsCount = SAPConnection.Children.Count

    'Not specified by user - connect to first session
    If sessionNo = -1 Then
        sessionNo = 1
    Else
        'Check if sessionNo is within boundaries (1 to number of sessions)
        If sessionNo < 1 Then
            sessionNo = 1
        End If
    End If

    'Check if session is busy - increase session number if it is
    If sessionNo <= sessionsCount Then
        For I = 1 To sessionsCount
            If sessionNo = I Then
                If SAPConnection.Children(I - 1).Busy Then
                    sessionNo = sessionNo + 1
                End If
            End If
        Next I
    End If
    
    'Create New Session if needed
    If sessionNo > sessionsCount Then
        Dim sessionCreated As Boolean
        
        sessionCreated = False
        
        'Wait till new Session will be created
        Do
            'It is impossible to create new session if session from which we are creating it is busy - we still have to wait (nooooo ;-/)
            If sessionCreated = False Then
                'Try to create session - hopefully one of currently opened sessions is not busy
                For I = 1 To sessionsCount
                    Set vSession = SAPConnection.Children(I - 1)
                    
                    If vSession.Busy = False Then
                        sessionCreated = True
                        vSession.CreateSession
                        Exit For
                    End If
                Next I
            Else
                If vSession.Busy = False Then
                    SID = SAP_GetWindowID(vSession, 1)
        
                    If SID <> "" Then
                        If vSession.FindByID(SID).Text = "Information" Then
                            If vSession.FindByID("wnd[1]/usr/txtMESSTXT1").Text = "Maximum number of sessions reached" Then
                                MsgBox "Maximum number of sessions reached", vbCritical, "SAP Initialization Error"
                                vSession.FindByID(SID).Close
                                GoTo Exit_Program
                            End If
                        End If
                    End If
                    
                    Application.Wait (Now + DateValue("00:00:01"))
                End If
            End If
            
            DoEvents
        Loop While SAPConnection.Children.Count < sessionsCount + 1
        
        'When new session is created - it is busy for a moment - wait while it is busy - we don't need any warnings (below)
        sessionNo = SAPConnection.Children.Count
        Set vSession = SAPConnection.Children(sessionNo - 1)
        
        While vSession.Busy
            DoEvents
        Wend
        
        'Try again to connect to specified session
        '(user might have created new session manually)
        GoTo New_Connection_Opened
    End If

    Err.Clear
    
    Set vSession = SAPConnection.Children(sessionNo - 1)

    If vSession.Busy Then
        MsgBox "Session is currently busy.", vbInformation, "SAP Initialization"
    End If
    
    If SAP_Activated(vSession, systemID) = False Then GoTo Exit_Program

    If invalidObject(vSession, "Error while initializing Session " & sessionNo) Then GoTo Error_Handler
    
    If vSession.Info.ScriptingModeReadOnly Then
        MsgBox "Server application has Scripting Mode set to READ ONLY. You will not be able to manipulate SAP objects !", vbCritical, "SAP Initialization Warning"
    End If
    
    If vSession.Info.IsLowSpeedConnection Then
        MsgBox "SAPGUI runs with low speed connection flag - scripting is very limited and might not work at all !", vbCritical, "SAP Initialization Warning"
    End If
    
Error_Handler:

    If Err.Number <> 0 Then
        'Permission denied - user has canceled action
        If Err.Number = 70 Then
            MsgBox "Conection to SAP canceled by User !", vbCritical, "Canceled by User"
        Else
            MsgBox "Connection to SAP failed ...", vbCritical, "SAP not available"
        End If
                
        Err.Clear
    Else
        If invalidObject(vSession, "Unknown error - Connection to SAP failed ...") = False Then
            '
        End If
    End If

Exit_Program:
    
    Erase listClients

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will 'check' if Session is active
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_Activated(vSession As Object, Optional systemID As String = "", Optional tCode As String = "") As Boolean
    Dim o As Object
    
    Dim isActive As Boolean
    Dim response As Variant
    
    Dim backupTestToolMode As Long

    SAP_Activated = False
    
    If vSession Is Nothing Then Exit Function
    If SAPApp Is Nothing Then Exit Function
    
CheckAgain:

    On Error Resume Next
    
    'Start Transaction - should raise an error if diconnected
    If tCode <> "" Then
        vSession.StartTransaction tCode
    End If
    
    If Err.Number = error_SAP_Disconnected Then
        Err.Clear
        response = MsgBox("Log in again, and click Yes to continue.", vbCritical + vbYesNo, "SAP got disconnected")

        If response = vbYes Then
            SAP_Init vSession, systemID

            GoTo CheckAgain
        End If
    End If

    If Err.Number = 0 Then
        SAP_Activated = True
    Else
        Err.Clear
    End If

    On Error GoTo -1
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will connect to specific sessionNo only
'   This allows me to connect to SAP session #2 while SAP session #1 is already connected to another macro running in parallel.
'   There are no safety checks here - only for super users :)))
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub SAP_InitSession(vSession As Object, sessionNo As Long)
    Set SAPGUIAuto = GetObject("SAPGUI")
    
    Set SAPApp = SAPGUIAuto.GetScriptingEngine()
    Set SAPConnection = SAPApp.Children(0)
    
    'SessionNo has to be converted to Integer
    Set vSession = SAPConnection.Children(CInt(sessionNo - 1))
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will navigate through menu bar
'   SAP_SelectMenu Session, "System > List > Save > Local file"
'   Default menuSeparator is '>' but you can change it if required
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_SelectMenu(vSession As Object, menuPath As String, Optional menuSeparator As String = ">") As Boolean
    Dim I As Long

    Dim listMenu() As String
    
    Dim o As Object
    Dim SID As String
    
    Dim menuExists As Boolean
    
    SAP_SelectMenu = False
    menuExists = False
    
    'Parse menu path
    ReDim listMenu(0): listMenu(0) = ""
    
    menuPath = UCase(menuPath)

    I = InStr(menuPath, menuSeparator)
    If I > 0 Then
        Do
            If listMenu(0) <> "" Then ReDim Preserve listMenu(UBound(listMenu) + 1)
            
            listMenu(UBound(listMenu)) = Trim(Mid(menuPath, 1, I - 1))
            
            menuPath = Mid(menuPath, I + 1)
            menuPath = Trim(menuPath)
            
            I = InStr(menuPath, ">")
        Loop While I > 0
    End If
    
    If menuPath <> "" Then
        If listMenu(0) <> "" Then ReDim Preserve listMenu(UBound(listMenu) + 1)
    
        listMenu(UBound(listMenu)) = Trim(menuPath)
    End If
    
    'Menu bar
    SID = "wnd[0]/mbar"
    
    'Loop through menu path
    For I = LBound(listMenu) To UBound(listMenu)
        menuExists = False
        
        'Check if our menu item exists (wildmatch)
        For Each o In vSession.FindByID(SID).Children
            If UCase(Trim(o.Text)) Like listMenu(I) Then
                menuExists = True
                
                'Update SAP ID (next time script will search only in submenu
                SID = o.ID: Set o = Nothing
                
                Exit For
            End If
        Next o
        
        'if menu exists --> continue loop
        If menuExists = False Then Exit For
    Next I
    
    If menuExists Then
        'If ChangeAble is True then menu is enabled and we can select it!
        If vSession.FindByID(SID).changeAble = True Then
            vSession.FindByID(SID).Select
        Else
            menuExists = False
        End If
    End If
    
    SAP_SelectMenu = menuExists
    
Exit_Program:

    Erase listMenu
    Set o = Nothing
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will convert file in 'Unconverted' format to format readable by Excel
'   SAP_FormatUnconvertedFile "Uncoverted file.TXT", "Formatted file.TXT"
'
'   - lines which are split after 1024 characters are concatenated back together
'   - empty rows '|------|' are replaced with vbCrLf
'   - all values in cells are 'trimmed' automatically
'   - if cells contain extra pipe '|', script replaces them with space !
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------

'2020-08-27
'Hmmm seems like in case there is a breakline @ 1024 character and another header cell is on a new line whole file format is broken
'Fixed with flagFixBrokenFormat ?
'                                      1023
'------------------------------------------
'                           |Partial BTLine
'|SCH No|Actual Shipment
'------------------------------------------

Function SAP_FormatUnconvertedFile(inputFileName As String, outpuFileName As String) As Boolean
    Dim I As Long
    Dim J As Long
    Dim K As Long
    
    Dim concatLen As Long
    Dim dataOffset As Long
    
    Dim fso As Object
    Dim inputFile As Object
    Dim outputFile As Object

    Dim s As String
    Dim fileLine As String
    
    Dim newLine As String
    Dim lastNewLine As String
    
    Dim flagWrite As Boolean
    
    Dim rowLen As Long
    
    Dim wordTrim As Variant
    
    Dim flagHeadersIdentified As Boolean
    Dim flagThisIsHeader As Boolean
    Dim flagFixBrokenFormat As Boolean
    Dim flagRemovePipes As Boolean

    Dim noOfColumns As Long
    Dim lastNoOfColumns As Long
    
    Dim rowSwitch As Long
    
    SAP_FormatUnconvertedFile = False
    
    rowSwitch = 0
    flagFixBrokenFormat = False

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set inputFile = fso.OpenTextFile(inputFileName, fsForReading)
    
    'On Error GoTo Error_Handler
    
    'Delete file if it exists already (we have to because of appending)
    FSO_DeleteFile outpuFileName
    
    'Appending method is lightning fast when compared to 'incremental' load to string variable and writing file afterwards
    'object.OpenTextFile (filename, [ iomode, [ create, [ format ]]])
    Set outputFile = fso.OpenTextFile(outpuFileName, fsForAppending, True)

    lastNoOfColumns = -1

    newLine = ""
    flagWrite = False
    flagHeadersIdentified = False

    Do Until inputFile.AtEndOfStream
        fileLine = inputFile.Readline
        
        s = Replace(fileLine, "-", "")
        If s = "" Or s = "||" Then
            fileLine = ""
        End If
        
        rowLen = Len(fileLine)
        
        '
        If flagFixBrokenFormat = False Then
            If rowLen = 1023 Then
                flagFixBrokenFormat = True
            End If
            
            If rowLen = 1024 Then
                flagFixBrokenFormat = True
            End If
            
            If flagFixBrokenFormat Then
                Application.statusBar = "SAP_FormatUnconvertedFile: file format is broken - working on it"
            End If
        End If

        '---
        If fileLine = "" Then fileLine = vbCrLf
        If fileLine = "|" Then fileLine = vbCrLf
        
        'It is safe to assume that we have identified headers - if both newLine lastLine start with pipe '|'
        If flagHeadersIdentified = False Then
            If Left(newLine, 1) = "|" Then
                If Left(fileLine, 1) = "|" Or fileLine = vbCrLf Then
                    newLine = newLine & "|" & fileLine
                    flagHeadersIdentified = True
                    flagThisIsHeader = True
                    flagWrite = True
                    lastNewLine = newLine
                    GoTo WriteHeader
                End If
            End If
        End If

ContinueWithData:

        'If fileLine <> vbCrLf Then
            newLine = newLine & fileLine
        'End If
        
        'Pipe is an indicator of new column --> row has to either begin with pipe '|' or dash '-'
        'script removes dashes '-'
        If Left(newLine, 1) = "|" Then
            If flagHeadersIdentified Then

WriteHeader:
                
                'Trim words
                wordTrim = Split(newLine, "|")
                noOfColumns = UBound(wordTrim)
                
                If flagFixBrokenFormat Then
                    If noOfColumns = lastNoOfColumns Then
                        flagWrite = True
                    End If
                Else
                    flagWrite = True
                End If
                
                If noOfColumns > lastNoOfColumns Then
                        flagWrite = flagWrite
                End If
                
                'Check data consistency - is number of columns different than last one ?
                If noOfColumns <> lastNoOfColumns Then
                    'Do we have an idea on how many columns there are ?
                    If lastNoOfColumns <> -1 Then
                        'We do ... let's check for extra pipes in data!
                        Dim wordLen() As String
                        
                        'Get length of all columns from previous row
                        wordTrim = Split(lastNewLine, "|")
                        
                        ReDim wordLen(UBound(wordTrim))
                        
                        For I = LBound(wordTrim) + 1 To UBound(wordTrim)
                            wordLen(I) = Len(wordTrim(I))
                        Next I
                        
                        'Check length of all columns in current row --> concat columns back if required (replace Pipe with space)
                        wordTrim = Split(newLine, "|")
                        
                        flagRemovePipes = False
                        
RestartLoop:
    
                        If UBound(wordTrim) > UBound(wordLen) Then
        
                            For I = LBound(wordTrim) + 1 To UBound(wordTrim) - 1
                                If I < UBound(wordLen) Then
                                    If Len(wordTrim(I)) <> wordLen(I) Then
                                        concatLen = 0
                                        K = 0
                                        '
                                        For J = I To UBound(wordTrim)
                                            concatLen = concatLen + Len(wordTrim(J))
                                            
                                            K = K + 1
                                            
                                            If K > 1 Then
                                                concatLen = concatLen + 1
                                            End If
                                            
                                            If concatLen = wordLen(I) Then
                                                Application.statusBar = "SAP_FormatUnconvertedFile: removing extra pipes '|'"
                                                flagRemovePipes = True
                                                
                                                For K = I + 1 To J
                                                    wordTrim(I) = wordTrim(I) & " " & wordTrim(K)
                                                Next K
                                                
                                                dataOffset = J - I
                                                
                                                For K = I + 1 To UBound(wordTrim) - dataOffset
                                                    wordTrim(K) = wordTrim(K + dataOffset)
                                                Next K
    
                                                ReDim Preserve wordTrim(UBound(wordTrim) - dataOffset)
                                                GoTo RestartLoop
    
                                            End If
                                        Next J
                                    End If
                                End If
                            Next I
                        End If

                        'Check one more time - after pipe removal
                        If flagRemovePipes Then
                            'wordTrim = Split(newLine, "|")
                            noOfColumns = UBound(wordTrim)
                            
                            If noOfColumns = lastNoOfColumns Then
                                flagWrite = True
                            Else
                                break
                            End If
                        End If
                            
                        'Keep old data for another data consisteny check
                        noOfColumns = lastNoOfColumns
                    End If
                End If
                    
                'Store previous data
                lastNoOfColumns = noOfColumns
            
                'Write to new file
                If flagWrite Then
                    'Rebuild newLine - trim
                    newLine = ""
    
                    '|Column1|Column2|Column3
                    For I = LBound(wordTrim) To UBound(wordTrim)
                        wordTrim(I) = Trim(wordTrim(I))
                        
                        If I > 0 Then
                            newLine = newLine & "|"
                        End If
                        
                        newLine = newLine & wordTrim(I)
                    Next I
                    
                    outputFile.Write vbCrLf
                    outputFile.Write newLine & "|"
                    
                    newLine = ""
                    flagWrite = False
                End If
                
                If flagThisIsHeader Then
                    flagThisIsHeader = False
                    'GoTo ContinueWithData
                End If
            End If
        Else
            outputFile.Write newLine
            newLine = ""
        End If
    Loop
    
    inputFile.Close
    
    outputFile.Close
    
Error_Handler:
    
    'If there is no error while writing to file - then we can assume everything was ok
    If Err.Number = 0 Then
        SAP_FormatUnconvertedFile = True
    Else
        Err.Clear
    End If
    
    On Error GoTo -1
    
    Application.statusBar = False

    Set inputFile = Nothing
    Set outputFile = Nothing
    Set fso = Nothing
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function is exporting data to excel. Exported format is determined based on file *extension*.
'
'   *** My recommendation ***
'   is to use .TXT, .XLSX (unconverted format) - it is fast, space-efficient, can handle huge files + data with lot of columns.
'       Spreadsheet format is causing issues if file contains columns having more than 1024 characters in 1 row - after 1024 characters there is a breakline messing up data!
'       MHTML format with >20 MB is not reliable, SAP might open new Excel instance and load file in it. (even with Application.IgnoreRemoteRequests = False) !
'
'   .TXT, .XLSX Unconverted - function exports file first to TXT format, then converts it, opens it in excel and finally saves file in XLSX format :)
'   .XLS Spreadsheet | Tab With tabs
'   Script will try to export:
'
'       1. By selecting menu:
'           1. System > List > Save > Local file
'           2. List > Export > Local file...
'       2. By searching for button with IconName B_DOWN
'       3. If ShellContainer is present by using context menu of ShellContainer
'
'   .MHTML function will try to export:
'
'       1. By selecting menu List > Export > Spreadsheet...
'       2. By searching for button with IconName LISVIE
'       3. If ShellContainer is present by using context menu of ShellContainer
'
'   WARNINGS:
'   MHTML export does not work reliably with bigger files (>20 MB)
'   MHTML export does not work with all GUIs :-/ (S4 works fine, Fusion does not work without user interaction in Save As dialog)
'
'   Usage:
'   SAP_ExportToExcel Session, "EXPORT.TXT", ["C:\USERS\ME\DESKTOP"] - will be saved as EXPORT.TXT, then converted and saved as EXPORT.XLSX!
'   SAP_ExportToExcel Session, "EXPORT.XLSX", ["C:\USERS\ME\DESKTOP"] - will be saved as EXPORT.TXT, then converted and saved as EXPORT.XLSX!
'   SAP_ExportToExcel Session, "EXPORT.XLS", ["C:\USERS\ME\DESKTOP"]
'   SAP_ExportToExcel Session, "EXPORT.MHTML", ["C:\USERS\ME\DESKTOP"]
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_ExportToExcel(vSession As Object, ByVal fileName As String, Optional filePath As String = "", Optional keepOpened As Boolean = False) As Boolean
    Const fileFormat_Unconverted = 0
    Const fileFormat_XLS As Long = 1
    Const fileFormat_MHTML As Long = 2
    
    Dim I As Long

    Dim o As Object
    Dim SID As String
    
    Dim exporting As Boolean
    Dim loopCounter As Byte
    
    Dim fileFormat As Long
    Dim iconName As String
    
    Dim fileExtension As String
    
    Dim WB As Workbook
    Dim backupIgnoreRemoteRequests As Boolean
    
    Dim newWorkbook As Boolean
    Dim listOpenedWorkbooks() As String
    
    ReDim listOpenedWorkbooks(0): listOpenedWorkbooks(0) = ""
    
    For Each WB In Application.Workbooks
        If listOpenedWorkbooks(0) <> "" Then ReDim Preserve listOpenedWorkbooks(UBound(listOpenedWorkbooks) + 1)
        
        listOpenedWorkbooks(UBound(listOpenedWorkbooks)) = WB.Path & "\" & WB.Name
    Next WB
    
    backupIgnoreRemoteRequests = Application.IgnoreRemoteRequests
    
    'Prevent opening files in new Excel instance (MHTML format)
    Application.IgnoreRemoteRequests = True
    
    loopCounter = 0
    exporting = False
    
    'Default filePath
    If filePath = "" Then filePath = filePathSaveAs
    
    'Update filePathSaveAs path
    SAP_SetFilePathSaveAs filePath
    
    filePath = filePathSaveAs
    
    'Safety check
    If SAP_Activated(vSession) = False Then GoTo Exit_Program
    
    'Safety check (required in case of asynchronous mode - why would we use it?)
    Do: Loop While vSession.Busy
    
    'Default name
    If fileName = "" Then fileName = "SAP_Export_" & Format(Now(), "YYYY-MM-DD HH MM SS") & ".XLS"
    
    fileExtension = getFileExtension(fileName)
    fileExtension = Trim(fileExtension)
    fileExtension = UCase(fileExtension)
    
    'Default - Local file (delimited by tabs)
    iconName = "B_DOWN"
    fileFormat = fileFormat_XLS
    
    'TODO: do we want to support more formats?
    Select Case fileExtension
    Case "TXT", "XLSX":
        fileFormat = fileFormat_Unconverted
        fileName = changeExtension(fileName, "TXT")
    
    Case "MHTML":
        iconName = "LISVIE"
        fileFormat = fileFormat_MHTML
        
    End Select
    
    'Navigate through menu
    If fileFormat = fileFormat_MHTML Then
        exporting = SAP_SelectMenu(vSession, "List>Export>Spreadsheet...")
    Else
        exporting = SAP_SelectMenu(vSession, "System>List>Save>Local file")
        If exporting = False Then exporting = SAP_SelectMenu(vSession, "List>Export>Local file...")
    End If
    
    'If downloading by menu was not possible ...
    If exporting = False Then
        'Search for download button in ToolBar
        If vSession.FindByID("wnd[0]/tbar[1]").Children.Count > 0 Then
            For Each o In vSession.FindByID("wnd[0]/tbar[1]").Children
                'Is this button which we are looking for?
                If o.iconName = iconName Then
                    'It is safer to work with SID than with object 'o' itself
                    SID = o.ID: Set o = Nothing
                    
                    vSession.FindByID(SID).Press
                    
                    exporting = True
                    
                    Exit For
                End If
            Next o
            
            Set o = Nothing
        End If

        If exporting = False Then
            'If shell container is available ...
            For Each o In vSession.FindByID("wnd[0]/usr").Children
                If UCase(SAP_GetSID(o.ID, "cntlCONTAINER")) = UCase("cntlCONTAINER") Then
                    
                    SID = "wnd[0]/usr/cntlCONTAINER/shellcont/shell"

                    Do
                        If fileFormat = fileFormat_MHTML Then
                            vSession.FindByID(SID).contextMenu
                            Do: Loop While vSession.Busy
                            
                            vSession.FindByID(SID).SelectContextMenuItem "&XXL"
                            Do: Loop While vSession.Busy
                        Else
                            vSession.FindByID(SID).PressToolbarContextButton "&MB_EXPORT"
                            Do: Loop While vSession.Busy
                            
                            vSession.FindByID(SID).SelectContextMenuItem "&PC"
                            Do: Loop While vSession.Busy
                        End If
                        
                        loopCounter = loopCounter + 1
                        
                        'There is a 'bug', in which conext menu is not working if Session is not maximized (?) ...
                        'So if macro unsuccessfully tried to export file through content menu for more than 10 times, we will activate window by minimizing & maximizing it
                        If loopCounter > 10 Then
                            vSession.FindByID("wnd[0]").Iconify
                            vSession.FindByID("wnd[0]").Maximize
                        End If
                        
                        'if we are not successful anyway, then I have no clue why it is not working! :)
                        If loopCounter > 20 Then Exit For
                    
                        'Wait for subwindow to appear
                    Loop While SAP_GetWindowID(vSession, 1) = ""
                    
                    exporting = True
                    
                    Exit For
                End If
            Next o
        
            Set o = Nothing
        End If
    End If
        
    'Specify format
    If exporting Then
        exporting = False
        
        'In case of MHTML format Radio buttons are children of wnd[1]/usr
        'In case of Local file (XLS) / Unconverted (TXT) format radio buttons are children of wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150
        
        'Find out which object ID is valid (we don't really need to use this function, but we can :))
        SID = SAP_GetValidObjectID(vSession, Array("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150", "wnd[1]/usr"), "wnd[1]/usr")
        
        If SID <> "" Then
            For Each o In vSession.FindByID(SID).Children
                If fileFormat = fileFormat_Unconverted Then
                    If UCase(Trim(o.Text)) = "UNCONVERTED" Then
                        exporting = True
                        Exit For
                    End If
                Else
                    If stringIsInArray(UCase(Trim(o.Text)), Array("SPREADSHEET", "TEXT WITH TABS", "EXCEL (IN MHTML FORMAT")) Then
                        exporting = True
                        Exit For
                    End If
                End If
            Next o
            
            If exporting = True Then
                SID = o.ID: Set o = Nothing
                
                vSession.FindByID(SID).Select
                
                'Ok, Continue
                vSession.FindByID("wnd[1]/tbar[0]/btn[0]").Press
            End If
            
            Set o = Nothing
        End If
    End If
    
    'Specify filepath and filename
    If exporting = True Then
        
        SID = SAP_GetWindowID(vSession, 1)
        
        If SID <> "" Then
            vSession.FindByID("wnd[1]/usr/ctxtDY_PATH").Text = filePath
            vSession.FindByID("wnd[1]/usr/ctxtDY_FILENAME").Text = fileName
            
            'Press Replace button
            exporting = False
            
            For Each o In vSession.FindByID("wnd[1]/tbar[0]").Children
                Select Case UCase(Trim(o.Text))
                Case "REPLACE":
                    SID = o.ID: Set o = Nothing
                    
                    vSession.FindByID(SID).Press
                    
                    exporting = True
                    
                    Exit For
                End Select
            Next o
    
            Set o = Nothing
        End If
    End If

    'If there is any subwindow available - then we MOST LIKELY received a warning message - thus export failed ...
    If exporting Then
        SID = SAP_GetWindowID(vSession, 1)
        
        If SID <> "" Then
            exporting = False
            Application.statusBar = "ERROR !"
            MsgBox "Error while saving file: " & fileName, vbOKOnly, "SAP EXPORT ERROR !"
        End If
    End If

    'Wait till exported (probably not required, but should not harm)
    Do: Loop While vSession.Busy
    
    If exporting Then
        'In case of unconverted format - we want to actually open file and convert it
        If fileFormat = fileFormat_Unconverted Then
            'Format file - save it to new temporary file
            If SAP_FormatUnconvertedFile(filePath & fileName, filePath & "$TEMP.FORMAT.UNCONVERTED." & fileName) Then
                
                'File extension has to be TXT !? Stupid excel ...
                Workbooks.OpenText fileName:=filePath & "$TEMP.FORMAT.UNCONVERTED." & fileName, Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", TrailingMinusNumbers:=True
                
                'TODO: is this reliable ?
                Set WB = ActiveWorkbook
                
                'Change Sheet name by default to Sheet1
                WB.ActiveSheet.Name = "Sheet1"
                WB.saveAs filePath & changeExtension(fileName, "XLSX"), fileFormat:=xlOpenXMLWorkbook
                
                'Delete temporary file
                FSO_DeleteFile filePath & "$TEMP.FORMAT.UNCONVERTED." & fileName
                
                'Update fileName
                fileName = changeExtension(fileName, "XLSX")
                
                'Close Workbook if we don't want to open it
                If keepOpened = False Then WB.Close
            End If
        End If
        
        'SAP opens MHTML file automatically
        If fileFormat = fileFormat_MHTML Then
            Dim innerTimer As Double
            
            innerTimer = Now
            
            loopCounter = 0
            
            Do
                'DoEvents should allow SAP to open exported file in active Excel instance
                DoEvents
                Application.Wait Now + TimeValue("00:00:01")
                
                'Was workbook opened in this Excel instance ?
                For Each WB In Application.Workbooks
                    If WB.Path & "\" & WB.Name = filePath & fileName Then
                        'Close Workbook if we don't want to open it
                        If keepOpened = False Then WB.Close
                        Exit Do
                    End If
                Next WB
                
                'In case of some GUIs we are not able to Save it automatically ... and only user can save file
                'That's why we have to check whether there was meanwhile opened new workbook (we cannot rely on user to save file to intended 'filePath' with 'fileName'
                If listOpenedWorkbooks(0) <> "" Then
                    For Each WB In Application.Workbooks
                        
                        newWorkbook = True
                        
                        For I = LBound(listOpenedWorkbooks) To UBound(listOpenedWorkbooks)
                            If listOpenedWorkbooks(I) = WB.Path & "\" & WB.Name Then
                                newWorkbook = False
                                Exit For
                            End If
                        Next I
                        
                        If newWorkbook = True Then Exit Do
                    Next WB
                End If
                
                loopCounter = loopCounter + 1
                
                'Force opening
                If loopCounter > 30 Then
                    If newWorkbook = False Then
                        Workbooks.Open filePath & fileName
                    End If
                End If
    
                'In order to prevent infinite loop we will wait only for 'exportTimeOut' seconds
                If Int((Now - innerTimer) * 60 * 60 * 24) > exportTimeOut Then
                    Exit Do
                End If
            Loop
        End If
    End If

Exit_Program:

    Set WB = Nothing

    Erase listOpenedWorkbooks

    Application.IgnoreRemoteRequests = backupIgnoreRemoteRequests
    
    SAP_ExportToExcel = exporting

    'Release objects from memory
    Set WB = Nothing
    Set o = Nothing
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Same as above, but will also open exported file
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_ExportToExcelAndOpen(vSession As Object, ByVal fileName As String, Optional filePath As String = "") As Boolean
    SAP_ExportToExcelAndOpen = False
    
    'Export
    If SAP_ExportToExcel(vSession, fileName, filePath, True) Then
        'Open
        Workbooks.Open (filePath & fileName)
    
        'Check if opened ...
        If ActiveWorkbook.Name = fileName Then SAP_ExportToExcelAndOpen = True
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will select variant
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_SelectVariant(vSession As Object, variantName As String) As Boolean
    Dim I As Long

    Dim o As Object
    
    Dim SID As String
    
    SAP_SelectVariant = False
    
    'Safety check
    If SAP_Activated(vSession) = False Then GoTo Exit_Program
        
    'Let's try it by using a SAP_SelectMenu function
    If SAP_SelectMenu(vSession, "Goto>Variants>Get...") = False Then
        
        'If it did not work thorugh menu, then search for a button with IconName B_VARI
        For Each o In vSession.FindByID("wnd[0]/tbar[1]").Children
            If o.iconName = "B_VARI" Then
                SID = o.ID: Set o = Nothing
                vSession.FindByID(SID).Press
                Exit For
            End If
        Next o
        
        Set o = Nothing
    End If
    
    'If we do have a subwindow available
    SID = SAP_GetWindowID(vSession, 1)
    
    If SID <> "" Then
        'If there is a subwindow with tile Find Variant, we have to specified which one ...
        If UCase(vSession.FindByID(SID).Text) Like UCase("*Find Variant*") Then
            
            vSession.FindByID("wnd[1]/usr/txtV-LOW").Text = variantName          'Variant
            vSession.FindByID("wnd[1]/usr/ctxtENVIR-LOW").Text = ""              'Environment
            vSession.FindByID("wnd[1]/usr/txtENAME-LOW").Text = ""               'Created by
            vSession.FindByID("wnd[1]/usr/txtAENAME-LOW").Text = ""              'Changed by
            vSession.FindByID("wnd[1]/usr/txtMLANGU-LOW").Text = ""              'Original language
            
            'Execute
            vSession.FindByID("wnd[1]/tbar[0]/btn[8]").Press
            
            SAP_SelectVariant = True
            
            'If there is a subwindow with title Information, than variant does not exist
            SID = SAP_GetWindowID(vSession, 1)

            If SID <> "" Then
                If UCase(vSession.FindByID(SID).Text) Like UCase("*Information*") Then
                    SAP_SelectVariant = False
                End If
            End If
            
        'If the subwindow does not has a title Find variant then we have to select variant by pressing it in a GridView
        Else
            SID = "wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell"
            
            'VARIANT
            For I = 1 To vSession.FindByID(SID).RowCount
                If vSession.FindByID(SID).GetCellValue(I - 1, "VARIANT") = variantName Then
                    If vSession.FindByID(SID).FirstVisibleRow + vSession.FindByID(SID).VisibleRowCount < I Then
                        vSession.FindByID(SID).FirstVisibleRow = I
                    End If
                    
                    'This has to be converted to integer
                    vSession.FindByID(SID).SelectedRows = CInt(I - 1)
                    
                    'Copy
                    vSession.FindByID("wnd[1]/tbar[0]/btn[2]").Press
                    
                    SAP_SelectVariant = True
                    
                    Exit For
                End If
            Next I
        End If
    End If

Exit_Program:
    
    Set o = Nothing
    
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Main function to start reports in either SQ00 or SQ01
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_StartSQ_Wrapper(vSession As Object, tCode As String, userGroup As String, queryName As String) As Boolean
    Dim I As Long
    
    Dim o As Object
    
    Dim SID As String
    Dim thisUserGroup As String
    Dim statusBar As String

    Dim queryExecuted As Boolean
    
    queryExecuted = False
    SAP_StartSQ_Wrapper = False
    statusBar = ""
    
    If SAP_Activated(vSession) = False Then GoTo Exit_Program
    
    queryName = UCase(Trim(queryName))
    userGroup = UCase(Trim(userGroup))
    
    vSession.StartTransaction (tCode)
    statusBar = SAP_GetStatusBar(vSession)
    
    If statusBar <> "" Then GoTo Exit_Program

    'Check if Query Area is Standard area
    If UCase(Trim(vSession.FindByID("wnd[0]/usr/txtRS38R-WSTEXT").Text)) Like UCase("*Standard Area*") Then
        queryExecuted = True
    Else
        'If not then select Standard area
        If SAP_SelectMenu(vSession, "Environment>Query Areas") = False Then GoTo Exit_Program
        
        For Each o In vSession.FindByID("wnd[1]/usr").Children
            If UCase(Trim(o.Text)) Like UCase("*Standard area*") Then
                SID = o.ID: Set o = Nothing
                
                vSession.FindByID(SID).Select
                vSession.FindByID("wnd[1]/tbar[0]/btn[2]").Press
                
                queryExecuted = True
                
                Exit For
            End If
        Next o
        
        Set o = Nothing
    End If
    
    'Check if current User Group is our userGroup
    thisUserGroup = SAP_GetTitleBar(vSession)
    
    'Query from User Group Z1-LOG: Initial Screen
    I = InStr(thisUserGroup, "Query from User Group")
    
    If I > 0 Then
        thisUserGroup = Mid(thisUserGroup, Len("Query from User Group") + 1)
        thisUserGroup = Mid(thisUserGroup, 1, InStr(thisUserGroup, ":") - 1)
        
        thisUserGroup = UCase(Trim(thisUserGroup))
    End If

    'Select User Group
    If thisUserGroup = userGroup Then
        queryExecuted = True
    Else
        'Press button Other user group
        If queryExecuted = True Then
            queryExecuted = False
                
            For Each o In vSession.FindByID("wnd[0]/tbar[1]").Children
                If UCase(Trim(o.ToolTip)) Like UCase("*Other user group*") Then
                    SID = o.ID: Set o = Nothing
                    
                    vSession.FindByID(SID).Press
                    
                    queryExecuted = True
                    
                    Exit For
                End If
            Next o
        End If

        'User Group Table
        If queryExecuted = True Then
            queryExecuted = False
            
            SID = "wnd[1]/usr/cntlGRID1/shellcont/shell"
            
            vSession.FindByID(SID).FirstVisibleRow = 0

            'msgbox vSession.FindByID(SID).CurrentCellColumn
            
            'Name - DBGBNUM
            For I = 0 To vSession.FindByID(SID).RowCount - 1
                If vSession.FindByID(SID).FirstVisibleRow + vSession.FindByID(SID).VisibleRowCount < I Then
                    vSession.FindByID(SID).FirstVisibleRow = I
                End If
                
                If UCase(Trim(vSession.FindByID(SID).GetCellValue(I, "DBGBNUM"))) = userGroup Then
                    
                    'This has to be converted to integer
                    vSession.FindByID(SID).SelectedRows = CInt(I)
                    vSession.FindByID("wnd[1]/tbar[0]/btn[0]").Press
                    
                    queryExecuted = True
                    
                    Exit For
                End If
            Next I
        End If
    End If
    
'Select Query
    If queryExecuted = True Then
        queryExecuted = False
        
        SID = "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell"
    
        vSession.FindByID(SID).FirstVisibleRow = 0

        'Name - QNUM
        For I = 0 To vSession.FindByID(SID).RowCount - 1
        
            If vSession.FindByID(SID).FirstVisibleRow + vSession.FindByID(SID).VisibleRowCount < I Then
                vSession.FindByID(SID).FirstVisibleRow = I
            End If
                            
            If UCase(Trim(vSession.FindByID(SID).GetCellValue(I, "QNUM"))) = queryName Then
                'This has to be converted to integer
                vSession.FindByID(SID).SelectedRows = CInt(I)
                
                queryExecuted = True
                
                Exit For
            End If
        Next I
    End If

    If queryExecuted = True Then
        'Execute query
        vSession.FindByID("wnd[0]/tbar[1]/btn[8]").Press

        statusBar = vSession.FindByID("wnd[0]/sbar").Text
        If statusBar <> "" Then queryExecuted = False
    End If

Exit_Program:

    SAP_StartSQ_Wrapper = queryExecuted

    If statusBar <> "" Then MsgBox statusBar, vbCritical, "SAP Start " & tCode

    Set o = Nothing
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will start a query queryName in transaction SQ00, in user group userGroup
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_StartSQ00(vSession As Object, userGroup As String, queryName As String) As Boolean
    SAP_StartSQ00 = SAP_StartSQ_Wrapper(vSession, "SQ00", userGroup, queryName)
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will start a query queryName in transaction SQ01, in user group userGroup
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_StartSQ01(vSession As Object, userGroup As String, queryName As String) As Boolean
    SAP_StartSQ01 = SAP_StartSQ_Wrapper(vSession, "SQ01", userGroup, queryName)
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will load all SAP GUI objects which are available in vSession (this might take a long time, especially transactions having huge amount of screen elements)
'   searchArea - specifiies where script should start loading
'   sortList - forces script to sort objects. Sorting slows down process!
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_LoadAllObjects(vSession As Object, Optional ByVal searchArea As String = "", Optional ByVal sortList As Boolean = False) As Boolean
    Dim I As Long
    Dim J As Long
    
    Dim o As Object
    
    Dim currentLevel As Long
    Dim currentPosition As Long
    
    Dim uniqueID As String
    
    Dim loadingIndicator As String
    
    SAP_LoadAllObjects = False
    
    loadingIndicator = "|"
    
    If SAP_Activated(vSession) Then
        
        'In order to speed up functions which are reading objects from same screen I am using this unique identifier (passportTransactionID & screen number & searchArea)
        'to find out whether list of object IDs should be refreshed
        'TODO: check if this unique ID is really unique :)
        uniqueID = vSession.passportTransactionID & vSession.Info.ScreenNumber & searchArea
        
        If uniqueID <> passportTransactionID Then
            passportTransactionID = uniqueID
        
            ReDim listAllSID(0)
        
            I = 0
            currentLevel = 0
            
            If searchArea <> "" Then
                listAllSID(currentLevel).ID = searchArea
            Else
                listAllSID(currentLevel).ID = "wnd[0]"
            End If
                
LoadDataFromNextLevel:
            
            If sortList Then currentPosition = currentLevel
            
            'Load all objects from searchArea
            If vSession.FindByID(listAllSID(currentLevel).ID).containerType = True Then
                
                If Not vSession.FindByID(listAllSID(currentLevel).ID).Children Is Nothing Then
                    If vSession.FindByID(listAllSID(currentLevel).ID).Children.Count > 0 Then
                        For Each o In vSession.FindByID(listAllSID(currentLevel).ID).Children
                            
                            ReDim Preserve listAllSID(UBound(listAllSID) + 1)
                            
                            '--- Object sorting - slows down process, so use only if you need it
                            
                            currentPosition = currentPosition + 1
    
                            If sortList Then
                                If UBound(listAllSID) > currentPosition Then
                                    For I = UBound(listAllSID) To currentPosition + 1 Step -1
                                        listAllSID(I) = listAllSID(I - 1)
                                    Next I
                                End If
                            End If
                            
                            '---
                            
                            listAllSID(currentPosition).ID = o.ID
                            listAllSID(currentPosition).typeValue = o.Type
                            listAllSID(currentPosition).textValue = o.Text
                            listAllSID(currentPosition).nameValue = o.Name
                            
                            'Safe extraction of changeAble property (not all objects have changeAble property)
                            listAllSID(currentPosition).changeAble = SAP_GetChangeAble(o)
                            
                            'If listAllSID(currentPosition).changeAble Then
                            '    Dim r As Long
                            '    r = GetRows(ThisWorkbook.Sheets("GUIObjectsUsage"), 1) + 1
                            '    ThisWorkbook.Sheets("GUIObjectsUsage").Cells(r, 1).Formula = o.type
                            'End If
    
                            listAllSID(currentPosition).containerType = o.containerType
                            
                            Set o = Nothing
                        Next o
    
                        Set o = Nothing
                    End If
                End If
            End If
                
            '|/-\
            Application.statusBar = "SAP: loading all objects (this might take a while) " & loadingIndicator
            
            If loadingIndicator = "|" Then
                loadingIndicator = "/"
            ElseIf loadingIndicator = "/" Then
                loadingIndicator = "-"
            ElseIf loadingIndicator = "-" Then
                loadingIndicator = "\"
            ElseIf loadingIndicator = "\" Then
                loadingIndicator = "|"
            End If
            
            If currentLevel < UBound(listAllSID) Then
                currentLevel = currentLevel + 1
                GoTo LoadDataFromNextLevel
            End If
        End If
    
        SAP_LoadAllObjects = True
    End If

    Application.statusBar = False

    Set o = Nothing
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will compare list of object IDs in variable (array) 'v' with all screen object IDs from vSession. It returns first valid object ID.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_GetValidObjectID(vSession As Object, ByVal v As Variant, Optional searchArea As String = "") As String
    Dim I As Long
    Dim J As Long
    
    Dim listSID() As String
    
    SAP_GetValidObjectID = vbNullString
    
    If IsArray(v) Then
        ReDim listSID(UBound(v))
        
        For I = LBound(v) To UBound(v)
            listSID(I) = v(I)
        Next I
    Else
        ReDim listSID(0)
        listSID(0) = v
    End If
    
    If SAP_LoadAllObjects(vSession, searchArea) Then
        For I = LBound(listSID) To UBound(listSID)
            For J = LBound(listAllSID) To UBound(listAllSID)
                If SAP_GetSID(listSID(I)) = SAP_GetSID(listAllSID(J).ID) Then
                    SAP_GetValidObjectID = listSID(I)
                    Exit Function
                End If
            Next J
        Next I
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function checks if ScrollBar needs to be moved in order to display next sreen
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_TableVScrollBarReadyToScroll(vSession As Object, ByVal rowNumber As Long, ByVal tableSID As String) As Boolean
    SAP_TableVScrollBarReadyToScroll = False
    
    If SAP_Activated(vSession) Then
        If rowNumber = (vSession.FindByID(tableSID).VerticalScrollBar.Position + vSession.FindByID(tableSID).VerticalScrollBar.PageSize) Then
            SAP_TableVScrollBarReadyToScroll = True
        End If
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function checks if ScrollBar needs to be moved in order to display next sreen
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_TableHScrollBarReadyToScroll(vSession As Object, ByVal colNumber As Long, ByVal tableSID As String) As Boolean
    SAP_TableHScrollBarReadyToScroll = False
    
    If SAP_Activated(vSession) Then
        If colNumber = (vSession.FindByID(tableSID).HorizontalScrollBar.Position + vSession.FindByID(tableSID).HorizontalScrollBar.PageSize) Then
            SAP_TableHScrollBarReadyToScroll = True
        End If
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function returns column name based on column title
'   for example, SAP Tree column name might be 'COL1' while column title is 'List fields'
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_TreeGetColumnNameByTitle(vTreeStructure As T_SAP_Tree, ByVal columnTitle) As String
    Dim I As Long
    
    SAP_TreeGetColumnNameByTitle = ""
    
    For I = LBound(vTreeStructure.columns) To UBound(vTreeStructure.columns)
        If columnTitle = vTreeStructure.columns(I).columnTitle Then
            SAP_TreeGetColumnNameByTitle = vTreeStructure.columns(I).columnName
            Exit Function
        End If
    Next I
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function returns column name based on column title (in case of SAP Tree column name might be 'COL1' while column title is 'List fields'
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_TreeGetColumnNumberByTitle(vTreeStructure As T_SAP_Tree, ByVal columnTitle) As String
    Dim I As Long
    
    SAP_TreeGetColumnNumberByTitle = -1
    
    For I = LBound(vTreeStructure.columns) To UBound(vTreeStructure.columns)
        If columnTitle = vTreeStructure.columns(I).columnTitle Then
            SAP_TreeGetColumnNumberByTitle = I
            Exit Function
        End If
    Next I
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function loads SAP tree structure
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_TreeLoadStructure(vSession As Object, ByVal vTreeStructureSID As String, vTreeStructure As T_SAP_Tree) As Boolean
    Dim I As Long

    Dim nodeKey As String
    Dim nodePath As String
    
    Dim countChildren As Long
    
    Dim o As Object
    
    Dim items() As String
    
    Dim loadingIndicator As String
    
    SAP_TreeLoadStructure = False
    
    nodePath = "1"
    
    loadingIndicator = "|"
    
    ReDim vTreeStructure.listTreeNodes(0)
    
    vTreeStructure.listTreeNodes(0).nodeKey = ""
    vTreeStructure.listTreeNodes(0).nodePath = ""
    vTreeStructure.listTreeNodes(0).nodeItems = Null

    vTreeStructure.SID = vTreeStructureSID
    
    'Get column names
    Set o = vSession.FindByID(vTreeStructureSID).GetColumnNames
    
    ReDim vTreeStructure.columns(o.length)
    
    For I = 0 To o.length - 1
        If CStr(o(I)) <> "" Then
            vTreeStructure.columns(I).columnName = CStr(o(I))
        End If
    Next I
    
    Set o = Nothing

    'Get column titles
    Set o = vSession.FindByID(vTreeStructureSID).GetColumnTitles
    
    For I = 0 To o.length - 1
        If CStr(o(I)) <> "" Then
            vTreeStructure.columns(I).columnTitle = CStr(o(I))
        End If
    Next I
    
    Set o = Nothing
    
    '---

    Do
        nodeKey = vSession.FindByID(vTreeStructureSID).GetNodeKeyByPath(nodePath)
        
        If nodeKey > "" Then
            If vTreeStructure.listTreeNodes(0).nodeKey <> "" Then ReDim Preserve vTreeStructure.listTreeNodes(UBound(vTreeStructure.listTreeNodes) + 1)
            
            '---
            
            vTreeStructure.listTreeNodes(UBound(vTreeStructure.listTreeNodes)).nodeKey = nodeKey
            vTreeStructure.listTreeNodes(UBound(vTreeStructure.listTreeNodes)).nodePath = nodePath
            
            ReDim items(0): items(0) = ""
            
            'Load all items
            For I = LBound(vTreeStructure.columns) To UBound(vTreeStructure.columns)
                If items(0) <> "" Then ReDim Preserve items(UBound(items) + 1)
                items(UBound(items)) = vSession.FindByID(vTreeStructureSID).GetItemText(nodeKey, vTreeStructure.columns(I).columnName)
            Next I
            
            vTreeStructure.listTreeNodes(UBound(vTreeStructure.listTreeNodes)).nodeItems = items
            
            '---
            
            countChildren = vSession.FindByID(vTreeStructureSID).GetNodeChildrenCount(nodeKey)
            
            If vSession.FindByID(vTreeStructureSID).IsFolderExpandable(nodeKey) Then countChildren = 1
            
            If countChildren > 0 Then
                'Expand nodes
                vSession.FindByID(vTreeStructureSID).ExpandNode vTreeStructure.listTreeNodes(UBound(vTreeStructure.listTreeNodes)).nodeKey
                nodePath = nodePath & "\" & 1
            Else
                nodePath = getNextNodePath(nodePath)
            End If
        
            '|/-\
            Application.statusBar = "SAP: loading tree structure (this might take a while) " & loadingIndicator
            
            If loadingIndicator = "|" Then
                loadingIndicator = "/"
            ElseIf loadingIndicator = "/" Then
                loadingIndicator = "-"
            ElseIf loadingIndicator = "-" Then
                loadingIndicator = "\"
            ElseIf loadingIndicator = "\" Then
                loadingIndicator = "|"
            End If
        Else
            If InStr(nodePath, "\") > 0 Then
                nodeKey = "do not exit please"
                nodePath = getNextNodePathParent(nodePath)
            End If
        End If
    Loop While nodeKey > ""
    
    Application.statusBar = False

    If vTreeStructure.listTreeNodes(0).nodeKey <> "" Then SAP_TreeLoadStructure = True
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   TODO: finish this
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub SAP_ExecuteInBackground(vSession As Object)
    If SAP_SelectMenu(vSession, "Program>Execute in Background") Then
        'Output device
        vSession.FindByID("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCL"
        
        'Continue
        vSession.FindByID("wnd[1]/tbar[0]/btn[13]").Press
        
        'Immediate
        vSession.FindByID("wnd[1]/usr/btnSOFORT_PUSH").Press
        
        'Save
        vSession.FindByID("wnd[1]/tbar[0]/btn[11]").Press
    End If
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   TODO: finish this
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_StartSA38(vSession As Object, programName As String) As Boolean
    SAP_StartSA38 = False
    
    vSession.StartTransaction ("SA38")
    vSession.FindByID("wnd[0]/usr/ctxtRS38M-PROGRAMM").Text = programName
    vSession.FindByID("wnd[0]/tbar[1]/btn[8]").Press
    
    SAP_StartSA38 = True
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   TODO: finish this
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_StartSQVI(vSession As Object, queryName As String) As Boolean
    Dim SID As String
    
    SAP_StartSQVI = False
    
    vSession.StartTransaction ("SQVI")
    vSession.FindByID("wnd[0]/usr/ctxtRS38R-QNUM").Text = queryName

    'Small improvement - executing via menu is faster than searching for object IDs
    If SAP.SAP_SelectMenu(vSession, "QuickView > Execute > Execute") Then
        SAP_StartSQVI = True
    Else
        'SQVI update - S4 button ID is      "wnd[0]/usr/btnP1"
        '            - Fusion button ID is  "wnd[0]/tbar[1]/btn[8]"
        SID = SAP_GetValidObjectID(vSession, Array("wnd[0]/tbar[1]/btn[8]", "wnd[0]/usr/btnP1"), "wnd[0]")
    
        If SID <> "" Then
            vSession.FindByID(SID).Press
            SAP_StartSQVI = True
        End If
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function SAP_SelectLayout(vSession As Object, layoutName As String) As Boolean
    Dim I As Long

    Dim o As Object
    
    Dim SID As String
    
    SAP_SelectLayout = False
    
    'Safety check
    If SAP_Activated(vSession) = False Then GoTo Exit_Program
        
    'Let's try it by using a SAP_SelectMenu function
    If SAP_SelectMenu(vSession, "Settings>Layout>Choose...") Then
        SID = "wnd[1]/usr/cntlGRID/shellcont/shell"
        
        'VARIANT
        For I = 1 To vSession.FindByID(SID).RowCount
            If vSession.FindByID(SID).GetCellValue(I - 1, "VARIANT") = layoutName Then
                If vSession.FindByID(SID).FirstVisibleRow + vSession.FindByID(SID).VisibleRowCount < I Then
                    vSession.FindByID(SID).FirstVisibleRow = I
                End If
                
                'This has to be converted to integer
                vSession.FindByID(SID).SelectedRows = CInt(I - 1)
                
                'Copy
                vSession.FindByID("wnd[1]/tbar[0]/btn[0]").Press
                
                SAP_SelectLayout = True
                
                Exit For
            End If
        Next I
    End If

Exit_Program:
    
    Set o = Nothing
    
End Function

'Function switches SAP ID based using SYSTEMID value
'SwitchSystem(Array("SID 1", "SID 2", "SID 3", ...))
'   --> returns SID 1 when SYSTEMID = 0
'   --> returns SID 2 when SYSTEMID = 1
'   --> returns SID 3 when SYSTEMID = 2

Function SwitchSystem(v As Variant) As String
    Dim I As Long
    If IsArray(v) Then
        For I = LBound(v) To UBound(v)
            If I = SAPSystemID Then
                SwitchSystem = v(I)
                Exit Function
            End If
        Next I
    End If
    
    SwitchSystem = CStr(v)
End Function

'Function selects date in date field (in Calendar)
Function SAP_SelectDate(vSession As Object, ByVal SID As String, ByVal strDate As String) As Boolean
    Dim I As Long

    Dim DD As String
    Dim MM As String
    Dim YYYY As String

    SAP_SelectDate = False
    
    DD = "": MM = "": YYYY = ""

    'Convert strDate do YYYYMMDD format
    '"20210802"
    
    'DD.MM.YYYY
    If strDate Like "*.*.*" Then
        I = InStr(strDate, ".")
        If I > 0 Then
            DD = Mid(strDate, 1, I - 1)
            strDate = Mid(strDate, I + 1)
        End If
    
        I = InStr(strDate, ".")
        If I > 0 Then
            MM = Mid(strDate, 1, I - 1)
            strDate = Mid(strDate, I + 1)
            
            YYYY = strDate
        End If
        
        strDate = YYYY & MM & DD
    End If

    'MM/DD/YYYY
    If strDate Like "*/*/*" Then
        I = InStr(strDate, "/")
        If I > 0 Then
            MM = Mid(strDate, 1, I - 1)
            strDate = Mid(strDate, I + 1)
        End If
    
        I = InStr(strDate, "/")
        If I > 0 Then
            DD = Mid(strDate, 1, I - 1)
            strDate = Mid(strDate, I + 1)
            
            YYYY = strDate
        End If
        
        strDate = YYYY & MM & DD
    End If

    'Set focus
    vSession.FindByID(SID).SetFocus
    
    'Possible entries
    vSession.FindByID("wnd[0]").SendVKey 4

    SID = SAP.SAP_GetWindowID(vSession, 1)

    If SID <> "" Then
        vSession.FindByID("wnd[1]/usr/cntlCONTAINER/shellcont/shell").focusDate = strDate
        vSession.FindByID("wnd[1]/usr/cntlCONTAINER/shellcont/shell").firstVisibleDate = "DAY_NAME"
        vSession.FindByID("wnd[1]/tbar[0]/btn[0]").Press
        SAP_SelectDate = True
    End If
End Function

'Function returns SAP ID of active tab for GuiTabStrip
'Input SID - SID of GuiTabStrip object
Function SAP_GetActiveTabSID(vSession As Object, SID As String) As String
    Dim o As Object

    SAP_GetActiveTabSID = ""

    If SID = "" Then Exit Function

    If Trim(vSession.FindByID(SID).Type) <> "GuiTabStrip" Then Exit Function

    'Loop through all tabs
    For Each o In vSession.FindByID(SID).Children
        'If object has any children - then it is most likely active tab
        'TODO: is this reliable ?
        If o.Children.length > 0 Then
            SAP_GetActiveTabSID = o.ID
            Exit For
        End If
    Next o

    Set o = Nothing
End Function

'Function selects tab by name in GuiTabStrip
'Input SID - SID of GuiTabStrip object
Function SAP_SelectTab(vSession As Object, SID As String, tabName As String) As Boolean
    Dim o As Object

    SAP_SelectTab = False
                
    If SID = "" Then Exit Function
                
    If Trim(vSession.FindByID(SID).Type) <> "GuiTabStrip" Then Exit Function
                
    'Loop through all tabs
    For Each o In vSession.FindByID(SID).Children
        'Check tab name
        If o.Text = tabName Then
            'Select tab
            o.Select
            
            SAP_SelectTab = True
            Exit For
        End If
    Next o
    
    Set o = Nothing
End Function
