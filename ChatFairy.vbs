Option Explicit

Const MAIN_LOOP_DELAY = 10
Const PLATFORM_X64 = False

' 时间单位
Const TIME_UNIT_MS    = 1
Const TIME_UNIT_S     = 1000
Const TIME_UNIT_M     = 60000
Const TIME_UNIT_H     = 3600000
Const TIME_UNIT_D     = 86400000

' 运行时错误吗
Const ERR_CODE_RUNTIME_UNSUPP_METHOD = 438

' 文件操作
Const ForReading         = 1  
Const ForWriting         = 2
Const ForAppending       = 8
Const TristateUseDefault = -2
Const TristateTrue       = -1
Const TristateFalse      = 0

' DynwrapX 
Const DYNWRAPX_REGISTER_TYPE_ALL       = 0
Const DYNWRAPX_REGISTER_TYPE_METHOD    = 1
Const DYNWRAPX_REGISTER_TYPE_ADDR      = 2
Const DYNWRAPX_REGISTER_TYPE_CODE      = 3
Const DYNWRAPX_REGISTER_TYPE_CALLBACK  = 4

' gdi32.dll
Const GDI32_WHITE_BRUSH  = 0
Const GDI32_TRANSPARENT  = 1
Const GDI32_COLOR_RED    = &H0000FF
Const GDI32_SRCCOPY      = &HCC0020
Const GDI32_WHITE_PEN    = &H000006
		
' CAPICOM           
Const CAPICOM_HASH_ALGORITHM_SHA1      = 0
Const CAPICOM_HASH_ALGORITHM_MD2       = 1
Const CAPICOM_HASH_ALGORITHM_MD4       = 2
Const CAPICOM_HASH_ALGORITHM_MD5       = 3
Const CAPICOM_HASH_ALGORITHM_SHA256    = 4
Const CAPICOM_HASH_ALGORITHM_SHA384    = 5
Const CAPICOM_HASH_ALGORITHM_SHA512    = 6

' WMPlayer
Const WMP_PLAY_STATE_UNDEFINED     = 0
Const WMP_PLAY_STATE_STOPPED       = 1
Const WMP_PLAY_STATE_PAUSED        = 2
Const WMP_PLAY_STATE_PLAYING       = 3
Const WMP_PLAY_STATE_SCAN_FORWARD  = 4
Const WMP_PLAY_STATE_SCAN_REVERSE  = 5
Const WMP_PLAY_STATE_BUFFERING     = 6
Const WMP_PLAY_STATE_WAITING       = 7
Const WMP_PLAY_STATE_MEDIA_ENDED   = 8
Const WMP_PLAY_STATE_TRANSITIONING = 9
Const WMP_PLAY_STATE_READY         = 10
Const WMP_PLAY_STATE_RECONNECTING  = 11
Const WMP_PLAY_STATE_LAST          = 12

' Event
Const EVENT_TYPE_MODE_CHANGE  = 1
Const EVENT_TYPE_DRAW_TEXT    = 2

' Tuling 消息分类
Const TULING_API_CODE_TEXT    = 100000
Const TULING_API_CODE_LINK    = 200000
Const TULING_API_CODE_NEWS    = 302000
Const TULING_API_CODE_COOK    = 308000
Const TULING_API_CODE_SONG    = 313000
Const TULING_API_CODE_POEM    = 314000

' Tuling 配置信息
Const STR_TULING_API_CONF_URL = "http://www.tuling123.com/openapi/api"
Const STR_TULING_API_CONF_KEY = "e2c9bfb2dfed425998546978ef712174"
Const STR_TULING_API_CONF_LOC = "成都"
Const STR_TULING_API_CONF_UID = "666"

' 所有模式
Const STR_WORKING_MODE_COMMAND       = "命令模式"
Const STR_WORKING_MODE_MUSIC         = "音乐模式"
Const STR_WORKING_MODE_TULING        = "聊天模式"

' 方法定义
Const STR_PLUGIN_METHOD_INIT         = "hasPlugin_Init"    ' Initialize
Const STR_PLUGIN_METHOD_TERM         = "hasPlugin_Term"    ' Terminate
Const STR_PLUGIN_METHOD_TIMER        = "hasPlugin_Timer"
Const STR_PLUGIN_METHOD_HANDLE       = "hasPlugin_Handle"

' 音乐模式
Const STR_PLUGIN_MUSIC_PLAYER_RANDOM = "随机播放"
Const STR_PLUGIN_MUSIC_PLAYER_PAUSE  = "暂停播放"
Const STR_PLUGIN_MUSIC_PLAYER_STOP   = "停止播放"
Const STR_PLUGIN_MUSIC_PLAYER_GO     = "继续播放"
Const STR_PLUGIN_MUSIC_PLAYER_NEXT   = "下一首"
Const STR_PLUGIN_MUSIC_PLAYER_PREV   = "上一首"
Const STR_PLUGIN_MUSIC_SCANNER_DIR   = "E:\Data\Kugou\夜的钢琴曲"

Dim g_objHtml
Dim g_objWshShell
Dim g_objFSO
Dim g_objDBHelper
Dim g_objPluginMgr
Dim g_objEventMgr
Dim g_objDynWrapX
Dim g_objCallbackDict
Dim g_strDBPath
Dim g_lngTimestamp
Dim g_intScreenW
Dim g_intScreenH

Set g_objHtml= CreateObject("htmlfile")
Set g_objWshShell = CreateObject("WScript.Shell")
Set g_objFSO = CreateObject("Scripting.FileSystemObject")
Set g_objDynWrapX = CreateObject("DynamicWrapperX")
Set g_objCallbackDict= Dict_NewDict()
Set g_objDBHelper = New DBHelper
Set g_objPluginMgr = New PluginMgr
Set g_objEventMgr = New EventMgr

g_strDBPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "data.db")
g_intScreenW = g_objHtml.parentWindow.screen.width
g_intScreenH = g_objHtml.parentWindow.screen.height

Call VBSMain()

Set g_objHtml = Nothing
Set g_objWshShell = Nothing 
Set g_objFSO = Nothing 
Set g_objDynWrapX = Nothing
Set g_objDBHelper = Nothing
Set g_objPluginMgr = Nothing
Set g_objEventMgr = Nothing
Set g_objCallbackDict = Nothing

Function VBSMain
	Dim objIgnoreList
	Dim strNewData
	Dim blnChange
	
	DBUtil_Connect g_objDBHelper, g_strDBPath
	Set objIgnoreList = Array_ToList(Array(strNewData, ",", "，", "", "？"))
	
	g_objPluginMgr.Register New DynWrapXPlugin      ' DynWrapX
	g_objPluginMgr.Register New DrawTextPlugin      ' 展示文本
	g_objPluginMgr.Register New CommandPlugin       ' 命令模式
	g_objPluginMgr.Register New MusicScannerPlugin  ' 音乐扫描
	g_objPluginMgr.Register New MusicPlayerPlugin   ' 音乐模式
	g_objPluginMgr.Register New TulingPlugin        ' 图灵机器人
	
	g_objPluginMgr.UpdateMode = STR_WORKING_MODE_MUSIC
	
	Do
		WScript.Sleep MAIN_LOOP_DELAY
		strNewData = Clipboard_GetData()
		blnChange = False 
		
		If Not objIgnoreList.Contains(strNewData) Then
			objIgnoreList.RemoveAt 0
			objIgnoreList.Insert 0, strNewData
			If g_objPluginMgr.HitMode(strNewData) Then 
				g_objPluginMgr.UpdateMode = strNewData
				g_objEventMgr.PostEvent EVENT_TYPE_MODE_CHANGE, strNewData
			Else
				g_objPluginMgr.UpdateKeyw = strNewData
			End If
			g_objEventMgr.PostEvent EVENT_TYPE_DRAW_TEXT, g_objPluginMgr.strMode _
				 & "：" & g_objPluginMgr.strKeyw
			blnChange = True 
		End If 
		
		g_objEventMgr.DispatchEvent
		
		If blnChange Then 
			g_objPluginMgr.DispatchHandle
		End If 
		
		g_lngTimestamp = g_lngTimestamp + MAIN_LOOP_DELAY
		g_objPluginMgr.DispatchTimer
	Loop
	
	Set objIgnoreList = Nothing 
End Function 

Class PluginMgr
	Public objPluginDict
	Public strMode
	Public strKeyw
	
	Public Sub Class_Initialize()
		Set objPluginDict = Dict_NewDict()
	End Sub 
	
	Public Sub Class_Terminate()
		Set objPluginDict = Nothing
	End Sub 
	
	Public Property Let UpdateMode(strMode)
		Me.strMode = strMode
	End Property
	
	Public Property Let UpdateKeyw(strKeyw)
		Me.strKeyw = strKeyw
	End Property 
	
	Public Function Update(strMode, strKeyw)
		Me.strMode = strMode
		Me.strKeyw = strKeyw
	End Function 
	
	Public Function HitMode(strText)
		Dim arrPlugin
		Dim objPlugin
		Dim blnHit
		
		arrPlugin = objPluginDict.Items
		For Each objPlugin In arrPlugin
			blnHit = objPlugin.ModeKeys.HitMode(strText)
			If blnHit Then
				Exit For 
			End If
		Next 
		HitMode = blnHit
		
		Set objPlugin = Nothing
		Erase arrPlugin
	End Function
	
	Public Function Register(objPlugin)
		Dim strPlugin
		
		strPlugin = TypeName(objPlugin)
		If objPluginDict.Exists(strPlugin) Then 
			Exit Function 
		End If 
		
		If PropertyExists(objPlugin, STR_PLUGIN_METHOD_INIT) Then 
			objPlugin.Plugin_Init
		End If
		
		objPluginDict.Add strPlugin, objPlugin
	End Function 
	
	Public Function Unregister(strPlugin)
		Dim objPlugin
		
		If Not objPluginDict.Exists(strPlugin) Then 
			Exit Function
		End If 
		
		Set objPlugin = objPluginDict(strPlugin)
		objPluginDict.Remove objPlugin
		
		If PropertyExists(objPlugin, STR_PLUGIN_METHOD_TERM) Then 
			objPlugin.Plugin_Term
		End If
		
		Set objPlugin = Nothing
	End Function 
	
	Public Function DispatchInit()
		Dim arrPlugin
		Dim objPlugin
		
		arrPlugin = objPluginDict.Items
		For Each objPlugin In arrPlugin
			If PropertyExists(objPlugin, STR_PLUGIN_METHOD_INIT) Then 
				objPlugin.Plugin_Init
			End If 
		Next
		
		Set objPlugin = Nothing
		Erase arrPlugin
	End Function 

	Public Function DispatchTimer
		Dim arrPlugin
		Dim objPlugin
		
		arrPlugin = objPluginDict.Items
		For Each objPlugin In arrPlugin
    		If PropertyExists(objPlugin, STR_PLUGIN_METHOD_TIMER) Then 
    			objPlugin.Plugin_Timer strMode, strKeyw
    		End If 
	    Next 
	    
	    Set objPlugin = Nothing 
	    Erase arrPlugin
	End Function 
	
	Public Function DispatchHandle
		Dim arrPlugin
		Dim objPlugin
		Dim blnHitMode
		Dim blnHitKeyw
		
		arrPlugin = objPluginDict.Items
		For Each objPlugin In arrPlugin
			blnHitMode = objPlugin.ModeKeys.HitMode(strMode)
			blnHitKeyw = objPlugin.ModeKeys.HitKeyw(strKeyw)
			If blnHitMode Then 
				Exit For 
			End If 
	    Next 
	    
	    If blnHitMode And blnHitKeyw Then 
	    	Debug.WriteLine "关键字命中，", strKeyw
			If PropertyExists(objPlugin, STR_PLUGIN_METHOD_HANDLE) Then 
				objPlugin.Plugin_Handle strMode, strKeyw
			End If 
	    End If 
	    
	    Set objPlugin = Nothing 
	    Erase arrPlugin
	End Function 
End Class 

Class EventMgr
	Public strDebugTag
	Public objMethodList
	Public objEventQueue
	
	Public Sub Class_Initialize()
		strDebugTag = "EventMgr："
		Set objMethodList = Collections_NewArrayList()
		Set objEventQueue = Collections_NewQueue()
	End Sub 
	
	Public Sub Class_Terminate()
		Set objMethodList = Nothing
		Set objEventQueue = Nothing
	End Sub 
	
	Public Function Register(objSubscriber, strMethod, strType)
		Dim objMethod
		
		Set objMethod = Dict_NewDict()
		objMethod.Add "object", objSubscriber
		objMethod.Add "method", strMethod
		objMethod.Add "type"  , strType
		objMethodList.Add objMethod
		
		Debug.WriteLine strDebugTag, TypeName(objSubscriber), " 注册事件 ",  strMethod, " 类型 ", strType
	End Function 
	
	Public Function Unregister(objSubscriber, strMethod, strType)
		Dim objMethod
		Dim objDelete
		Dim varCompMethod
		Dim varCompType
		Dim intCount
		
		For Each objMethod In objMethodList
			If objMethod("object") Is objSubscriber Then 
				intCount = intCount + 1
			End If 
			If objMethod("method") = strMethod Or IsNull(strMethod) Then 
				intCount = intCount + 1
			End If 
			If objMethod("type") = strType Or IsNull(strType) Then 
				intCount = intCount + 1
			End If 
			If intCount = 3 Then 
				Set objDelete = objMethod
				Exit For 
			End If 
		Next 
		objMethodList.Remove objDelete
		
		Set objMethod = Nothing
		Set objDelete = Nothing
	End Function 
	
	Public Function PostEvent(strType, varEvent)
		Dim objEvent
		
		Set objEvent = Dict_NewDict()
		objEvent.Add "type", strType
		objEvent.Add "data", varEvent
		objEventQueue.Enqueue objEvent
		
		Set objEvent = Nothing
	End Function 

	Public Function DispatchEvent
		Dim objEvent
		Dim objMethod
		Dim objDst
		Dim strFunc
		
		While objEventQueue.Count > 0
			Set objEvent = objEventQueue.Dequeue
			For Each objMethod In objMethodList
				If  objMethod("type") = objEvent("type") Then 
					Set objDst = objMethod("object")
					strFunc = objMethod("method")
					Execute "objDst." & strFunc & "(objEvent)"
				End If 
			Next 
		Wend 
		
		Set objEvent = Nothing
		Set objMethod = Nothing 
		Set objDst = Nothing
	End Function 
End Class 

Class DynWrapXPlugin
	Public strDebugTag
	Public hasPlugin_Init
	Public objModeKeys
	Public objService
	
	Private Sub Class_Initialize()
		strDebugTag = "DynWrapXPlugin："
		Set objModeKeys = New ModeKeysCls
		Set objService = New DynWrapXService
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
		Set objService = Nothing
	End Sub 
	
	Public Property Get ModeKeys()
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function Plugin_Init()
		Dim objList
		Dim objDict
		Dim strBuffer
		Dim strMethod
		Dim strConvention
		Dim strParams
		Dim strReturn
		
		Debug.WriteLine strDebugTag, "Init"
		g_objCallbackDict.RemoveAll
		
		Set objList = objService.QueryByType(DYNWRAPX_REGISTER_TYPE_METHOD)
		For Each objDict In objList
			strBuffer = objDict("library")
			If Len(objDict("badname")) > 0 Then 
				strBuffer = strBuffer & ":" & objDict("badname")
			ElseIf Len(objDict("ordinal")) > 0 Then 
				strBuffer = strBuffer & ":" & objDict("ordinal")
			End If 
			strMethod = objDict("method")
			strConvention = "f=" & objDict("convention")
			strParams = "i=" & objDict("params")
			strReturn = "r=" & objDict("return")
			Register "Register", strBuffer, strMethod, strConvention, strParams, strReturn
		Next 
		
		Set objList = objService.QueryByType(DYNWRAPX_REGISTER_TYPE_ADDR)
		For Each objDict In objList
			strBuffer = objDict("addr")
			strMethod = objDict("method")
			strConvention = "f=" & objDict("convention")
			strParams = "i=" & objDict("params")
			strReturn = "r=" & objDict("return")
			Register "RegisterAddr", strBuffer, strMethod, strConvention, strParams, strReturn
		Next 
		
		Set objList = objService.QueryByType(DYNWRAPX_REGISTER_TYPE_CODE)
		For Each objDict In objList
			strBuffer = objDict("code")
			strMethod = objDict("method")
			strConvention = "f=" & objDict("convention")
			strParams = "i=" & objDict("params")
			strReturn = "r=" & objDict("return")
			Register "RegisterCode", strBuffer, strMethod, strConvention, strParams, strReturn
		Next 
		
		Set objList = objService.QueryByType(DYNWRAPX_REGISTER_TYPE_CALLBACK)
		For Each objDict In objList
			strBuffer = objDict("callback")
			strConvention = "f=" & objDict("convention")
			strParams = "i=" & objDict("params")
			strReturn = "r=" & objDict("return")
			Register "RegisterCallback", strBuffer, "", "f=", strParams, strReturn
		Next 
		
		Set objList = Nothing
		Set objDict = Nothing
	End Function 
	
	Function Register(strRegister, strLibrary, strMethod, strConvention, strParams, strReturn)
		Dim strStat
		Dim ptrFunc
		
		strLibrary = """" & strLibrary & """"
		If strRegister = "RegisterCallback" Then
			strLibrary = "(GetRef(" & strLibrary & ")"
		End If 
		
		strStat = "g_objDynWrapX." & strRegister & " " & strLibrary & ", "
		
		If strMethod <> "" Then 
			strMethod = """" & strMethod & """"
			strStat = strStat & strMethod & ", "
		End If 
		
		If strConvention <> "f=" Then 
			strConvention = """" & strConvention & """"
			strStat = strStat & strConvention & ", "
		End If 
		
		If strParams <> "i=" Then 
			strParams = """" & strParams & """"
			strStat = strStat & strParams & ", "
		End If 
		
		If strReturn <> "r=" Then 
			strReturn = """" & strReturn & """"
			strStat = strStat & strReturn & ", "
		End If 
		
		strStat = Left(strStat, Len(strStat)-2)

		If strRegister = "RegisterCallback" Then 
			strStat = strStat & ")"
			Debug.WriteLine strDebugTag, strStat
			ptrFunc = Eval(strStat)
			g_objCallbackDict.Add strLibrary, ptrFunc
		Else 
			Debug.WriteLine strDebugTag, strStat
			Execute strStat
		End If 
	End Function 
End Class 

Class DrawTextPlugin
	Public strDebugTag
	Public hasPlugin_Init
	Public objModeKeys
	Public objDrawText
	Public strText
	Public strBuff
	Public intX, intY, intW, intH
	Public intBuff
	Public intIndex
	
	Private Sub Class_Initialize()
		strDebugTag = "DrawTextPlugin："
		Set objModeKeys = New ModeKeysCls
		Set objDrawText = CreateObject("Tiky.ChatFairy.CFDrawText")
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
		Set objDrawText = Nothing
	End Sub 
	
	Public Property Get ModeKeys()
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function OnDrawTextEvent(objEvent)
		Debug.WriteLine strDebugTag, "OnDrawTextEvent, Data, ", objEvent("data")
		
		objDrawText.UpdateText(objEvent("data"))
	End Function 
	
	Public Function Plugin_Init()
		Debug.WriteLine strDebugTag, "Init"
		
		objDrawText.CreateWindowT g_intScreenW-300, g_intScreenH-100, 300, 20
		objDrawText.UpdateText ""
		objDrawText.UpdateTextColor RGB(255, 0, 0)
		
		g_objEventMgr.Register Me, "OnDrawTextEvent", EVENT_TYPE_DRAW_TEXT
	End Function 
End Class 

Class CommandPlugin
	Public strDebugTag
	Public hasPlugin_Init
	Public hasPlugin_Handle
	Public objModeKeys
	Public objService
	
	Private Sub Class_Initialize()
		strDebugTag = "CommandPlugin："
		Set objModeKeys = New ModeKeysCls
		Set objService = New CommandService
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
		Set objService = Nothing
	End Sub
	
	Public Property Get ModeKeys
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function Plugin_Init()
		Debug.WriteLine strDebugTag, "Init"
		
		objModeKeys.RegMode STR_WORKING_MODE_COMMAND
		objModeKeys.RegKeywAll True
	End Function 
	
	Public Function Plugin_Handle(strMode, strKeyw)
		Dim objCmdList
		Dim objCmdDict
		Dim strCmdText
		Dim strPerformer
		Dim strPerformArgs
		Dim strPerformObject
		
		strKeyw = DBUtil_Filter(strKeyw)
		Debug.WriteLine strDebugTag, "Handle, ", strMode, ", ", strKeyw
		Set objCmdList = objService.QueryByKeyword(strKeyw)
		
		If objCmdList.Count = 0 Then 
			Exit Function 
		End If 
		
		Set objCmdDict = objCmdList.Item(0)
		strPerformer = objCmdDict("performer")
		strPerformArgs = objCmdDict("performArgs")
		strPerformObject = objCmdDict("performObject")
		
		strCmdText = strPerformer
		
		If Len(strPerformObject) > 0 Then 
			strCmdText = strCmdText & Space(1)
			strCmdText = strCmdText & strPerformObject
		End If 
		
		If Len(strPerformArgs) > 0 Then 
			strCmdText = strCmdText & Space(1)
			strCmdText = strCmdText & strPerformArgs
		End If 
		
		Debug.WriteLine strDebugTag, "执行 ", strCmdText
		
		Select Case objCmdDict("category")
			Case 1
				g_objWshShell.Run strCmdText
			Case 2
				ExecuteGlobal strCmdText
			Case Else 
		End Select 
	End Function 
End Class 

Class MusicScannerPlugin
	Public strDebugTag
	Public hasPlugin_Init
	Public objModeKeys
	Public objService
	Public objDirList
	
	Private Sub Class_Initialize()
		strDebugTag = "MusicScannerPlugin："
		Set objModeKeys = New ModeKeysCls
		Set objService = New MusicService
		Set objDirList = Collections_NewArrayList
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
		Set objService = Nothing
	End Sub 
	
	Public Property Get ModeKeys()
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function Plugin_Init()
		Dim objList
		Dim objTemp
		Dim strFile
		Dim strID
		
		Debug.WriteLine strDebugTag, "Init"
		objDirList.Add STR_PLUGIN_MUSIC_SCANNER_DIR
		
		Set objList = objService.Query()
		For Each objTemp In objList 
			RefreshRecord objTemp
		Next
		
		Set objList = objService.Query()
		For Each objTemp In objDirList
			InsertRecord ScanDirectory(objTemp), objList
		Next 
		
		Set objList = Nothing 
		Set objTemp = Nothing
	End Function 
	
	Public Function RefreshRecord(objRecord)
		Dim strID
		Dim strFile
		
		strID   = objRecord("id")
		strFile = objRecord("path")
		
		If Not g_objFSO.FileExists(strFile) Then
			objService.DeleteByID strID
			Exit Function 
		End If
	End Function
	
	Public Function ScanDirectory(strDir)
		Dim objFolder
		Dim objFiles
		Dim objFile
		Dim objList
		Dim objDict
		Dim strExt
		
		Set objFolder = g_objFSO.GetFolder(strDir)
		Set objFiles = objFolder.Files
		Set objList = Collections_NewArrayList()
		
		For Each objFile in objFiles
			strExt = UCase(g_objFSO.GetExtensionName(objFile.path))
			If  strExt = "MP3" Or strExt = "MP4" Then
				Set objDict = Dict_NewDict()
				objDict.Add "name", split(objFile.Name, ".")(0) 
				objDict.Add "path", objFile.Path
				objDict.Add "type", strExt
				objDict.Add "size", objFile.Size
				objDict.Add "offline", "1"
				objList.Add objDict
			End If 
		Next
		
		Set ScanDirectory = objList
		
		Set objFolder = Nothing
		Set objFiles = Nothing
		Set objFile = Nothing
		Set objList = Nothing
		Set objDict = Nothing 
	End Function 
	
	Public Function InsertRecord(objScanList, objReferList)
		Dim objDict
		Dim objTemp
		Dim arrItem
		Dim strTemp
		
		Set objDict = Dict_NewDict()
		
		For Each objTemp In objReferList
			strTemp = objTemp("path")
			If Not objDict.Exists(strTemp) Then 
				objDict.Add strTemp, objTemp
			End If 
		Next 
		For Each objTemp In objScanList
			strTemp = objTemp("path")
			If Not objDict.Exists(strTemp) Then 
				objDict.Add strTemp, objTemp
			End If
		Next
		For Each objTemp In objReferList
			strTemp = objTemp("path")
			objDict.Remove strTemp
		Next 
		
		arrItem = objDict.Items
		For Each objTemp In arrItem
			objService.Insert objTemp
		Next 
		
		Set objDict = Nothing
		Set objTemp = Nothing
	End Function
End Class 

Class MusicPlayerPlugin
	Public strDebugTag
	Public hasPlugin_Init
	Public hasPlugin_Timer
	Public hasPlugin_Handle
	Public objModeKeys
	Public objService
	Public objPlayer
	Public objDict
	Public strKeyw
	
	Private Sub Class_Initialize()
		strDebugTag = "MusicPlayerPlugin："
		Set objModeKeys = New ModeKeysCls
		Set objService = New MusicService
		Set objPlayer = CreateObject("WMPlayer.OCX.7")
		Set objDict = Dict_NewDict()
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
		Set objService = Nothing
		Set objPlayer = Nothing
		Set objDict = Nothing 
	End Sub
	
	Public Property Get ModeKeys()
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function OnModeChangeEvent(objEvent)
		Dim strMode
		
		Debug.Write strDebugTag, "OnModeChangeEvent, "
		strMode = objEvent("data")
		If strMode <> STR_WORKING_MODE_MUSIC Then 
			Debug.WriteLine "音乐播放停止"
			objPlayer.Controls.Stop
		End If 
	End Function 
	
	Public Function Plugin_Init()
		Dim objList
		Dim objTemp
		
		Debug.WriteLine strDebugTag, "Init"
		
		objModeKeys.RegMode STR_WORKING_MODE_MUSIC
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_RANDOM
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_PAUSE
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_STOP
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_GO
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_NEXT
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_PREV
		
		Set objList = objService.Query()
		For Each objTemp In objList 
			objDict.Add objTemp("id"), objTemp("path")
		Next
		
		Debug.WriteLine strDebugTag, "读取歌曲完成，共 ", objList.Count, " 首 "
		g_objEventMgr.Register Me, "OnModeChangeEvent", EVENT_TYPE_MODE_CHANGE
		g_objEventMgr.PostEvent EVENT_TYPE_DRAW_TEXT, "音乐插件：读取歌曲 " & objList.Count & " 首"
	End Function 
	
	Public Function Plugin_Handle(strMode, strKeyw)
		Debug.WriteLine strDebugTag, "Handle, ", strMode, ", ", strKeyw
		Select Case strKeyw
			Case STR_PLUGIN_MUSIC_PLAYER_RANDOM
				RandomPlay
			Case STR_PLUGIN_MUSIC_PLAYER_PAUSE
				objPlayer.Controls.Pause
			Case STR_PLUGIN_MUSIC_PLAYER_STOP
				objPlayer.Controls.Stop
			Case STR_PLUGIN_MUSIC_PLAYER_GO
				objPlayer.Controls.Play
			Case STR_PLUGIN_MUSIC_PLAYER_NEXT
				NormalPlay
			Case STR_PLUGIN_MUSIC_PLAYER_PREV
				ReversePlay
			Case Else
		End Select 
		Me.strKeyw = strKeyw
	End Function 
	
	Public Function Plugin_Timer(strMode, strKeyw)
		If g_lngTimestamp Mod TIME_UNIT_S <> 0 Then 
			Exit Function
		End If
		
		If Not objModeKeys.HitMode(strMode) Then 
			Exit Function
		End If 
		
		If Not objModeKeys.HitKeyw(Me.strKeyw) Then 
			Exit Function
		End If 
	
		g_objEventMgr.PostEvent EVENT_TYPE_DRAW_TEXT, strMode & "：" & objPlayer.status
		
		If objPlayer.PlayState = WMP_PLAY_STATE_PLAYING Then
			Exit Function
		End If
		
		Select Case strKeyw
			Case STR_PLUGIN_MUSIC_PLAYER_RANDOM
				RandomPlay
			Case STR_PLUGIN_MUSIC_PLAYER_PAUSE
				
			Case STR_PLUGIN_MUSIC_PLAYER_STOP
				
			Case STR_PLUGIN_MUSIC_PLAYER_GO
				NormalPlay
			Case STR_PLUGIN_MUSIC_PLAYER_NEXT
				NormalPlay
			Case STR_PLUGIN_MUSIC_PLAYER_PREV
				ReversePlay
			Case Else
		End Select 
	End Function 
	
	Public Function RandomPlay()
		Dim arrFile
		Dim intRnd
		
		arrFile = objDict.Items
		intRnd = Math_Rnd(0, objDict.Count-1)
		
		objPlayer.Url = arrFile(intRnd)
	End Function 
	
	Public Function NormalPlay()
		Dim arrFile
		Dim intIndex
		
		arrFile = objDict.Items
		intIndex = Array_FirstIndexOf(arrFile, objPlayer.Url)
		
		If intIndex >= UBound(arrFile) Then
			intIndex = 0
		Else 
			intIndex = intIndex + 1
		End If 
		
		objPlayer.Url = arrFile(intIndex)
	End Function
	
	Public Function ReversePlay()
		Dim arrFile
		Dim intIndex
		
		arrFile = objDict.Items
		intIndex = Array_FirstIndexOf(arrFile, objPlayer.Url)
		
		If intIndex <= 0 Then
			intIndex = UBound(arrFile)
		Else 
			intIndex = intIndex - 1
		End If 
		
		objPlayer.Url = arrFile(intIndex)
	End Function 
End Class 

Class TulingPlugin
	Public strDebugTag
	Public hasPlugin_Init
	Public hasPlugin_Handle
	Public objModeKeys
	
	Private Sub Class_Initialize()
		Dim arrKeyw
		
		strDebugTag = "TulingPlugin："
		Set objModeKeys = New ModeKeysCls
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
	End Sub 
	
	Public Property Get ModeKeys()
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function Plugin_Init()
		Debug.WriteLine strDebugTag, "Init"
		
		objModeKeys.RegMode STR_WORKING_MODE_TULING
		objModeKeys.RegKeywAll True
	End Function 
	
	Public Function Plugin_Handle(strMode, strKeyw)
		Dim objDict
		Dim objList
		
		If objModeKeys.HitMode(strKeyw) Then 
			Exit Function
		End If 
		
		Set objDict = HTTP_Tuling_Ask(strKeyw)
		Debug.WriteLine strDebugTag, "Handle, ", strMode, ", ", strKeyw
		Debug.WriteLine strDebugTag, objDict("text")
		
		Select Case objDict("code")
			Case TULING_API_CODE_TEXT 
				' 说话
			Case TULING_API_CODE_LINK 
				g_objWshShell.Run objDict("url")
			Case TULING_API_CODE_NEWS, TULING_API_CODE_COOK 
				Set objList = objDict("list")
				g_objWshShell.Run objList(Math_Rnd(0, objList.Count))("detailurl")
		End Select 
		
		SP_VOICE_Speak objDict("text")
		
		Set objDict = Nothing
		Set objList = Nothing
	End Function 
End Class 

' 不要New，方便写新模块时使用
Class SamplesPlugin
	Public strDebugTag
	Public hasPlugin_Init
	Public hasPlugin_Timer
	Public hasPlugin_Handle
	Public objModeKeys
	
	Private Sub Class_Initialize()
		strDebugTag = "SamplesPlugin："
		Set objModeKeys = New ModeKeysCls
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
	End Sub 
	
	Public Property Get ModeKeys()
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function Plugin_Init()
		Debug.WriteLine strDebugTag, "Init"
		objModeKeys.RegMode "例子模式"
		objModeKeys.RegKeywAll True
	End Function 
	
	Public Function Plugin_Handle(strMode, strKeyw)
		Debug.WriteLine strDebugTag, "Handle, ", strMode, ", ", strKeyw
	End Function 
	
	Public Function Plugin_Timer(strMode, strKeyw)
		
	End Function 
End Class 

Class ModeKeysCls
	Public objMode
	Public objKeys
	Public blnAcceptAll
	
	Private Sub Class_Initialize()
		Set objMode = Collections_NewArrayList()
		Set objKeys = Collections_NewArrayList()
		blnAcceptAll = False
	End Sub 
	
	Private Sub Class_Terminate()
		Set objMode = Nothing 
		Set objKeys = Nothing 
	End Sub 
	
	Public Function RegMode(strMode)
		objMode.Add strMode
	End Function 
	
	Public Function RegKeyw(strKeyw)
		objKeys.Add strKeyw
	End Function 
	
	Public Function RegModeFromArray(arrMode)
		Dim strMode
		
		For Each strMode In arrMode
			RegMode(strMode)
		Next 
	End Function 
	
	Public Function RegKeywFromArray(arrKeyw)
		Dim strKeyw
		
		For Each strKeyw In arrKeyw
			RegKeyw(strKeyw)
		Next 
	End Function 
	
	Public Function RegKeywAll(blnAcceptAll)
		Me.blnAcceptAll = blnAcceptAll
	End Function 
	
	Public Function HitMode(strMode)
		Dim blnHit
		blnHit = objMode.Contains(strMode)
		HitMode = blnHit
	End Function
	
	Public Function HitKeyw(strKeyw)
		Dim blnHit
		
		If blnAcceptAll Then 
			blnHit = True
		Else
			blnHit = objKeys.Contains(strKeyw)
		End If 
		HitKeyw = blnHit
	End Function 
	
	Public Function HitModeOrKeyw(strModeOrKeyw)
		Dim blnHitMode
		Dim blnHitKeyw
		
		blnHitMode = HitMode(strModeOrKeyw)
		blnHitKeyw = HitKeyw(strModeOrKeyw)
		HitModeOrKeyw = blnHitMode Or blnHitKeyw
	End Function 
End Class 

Class DynWrapXService
	Public strDebugTag

	Private Sub Class_Initialize()
		strDebugTag = "DynWrapXService："
	End Sub 
	
	Private Sub Class_Terminate()

	End Sub
	
	Public Function QueryByType(intType)
		Dim strSQL
		
		Select Case intType
			Case 0 ' all
				strSQL = "select * from dynwrapx where enable=1;"
			Case 1 ' func
				strSQL = "select library,badname,ordinal," 
			Case 2 ' addr
				strSQL = "select addr,"
			Case 3 ' code
				strSQL = "select code,"
			Case 4 ' callback
				strSQL = "select callback,"
			Case Else 
				Exit Function
		End Select 

		If intType <> 0 Then 
			strSQL = strSQL & "id,method,convention,params,return,enable,type,createTime from dynwrapx where type='" 
			strSQL = strSQL & intType
			strSQL = strSQL & "' and enable=1;"
		End If 
		
		Set QueryByType = DBUtil_QueryByNativeSQL(g_objDBHelper, strSQL)
	End Function
	
	Public Function QueryByMethod(strMethod)
		Dim strSQL
		
		strSQL = "select * from dynwrapx "
		strSQL = strSQL & "where method='"
		strSQL = strSQL & strMethod 
		strSQL = strSQL & "' and enable=1;"
		
		Set QueryByMethod = DBUtil_QueryByNativeSQL(g_objDBHelper, strSQL)
	End Function 
	
	Public Function Insert(objDict)
		Dim strCols
		Dim strVals
		Dim strSQL
		
		strCols = DBUtil_GenerateStringByArray(objDict.Keys)
		strVals = DBUtil_GenerateStringByArray(objDict.Items)
		
		strSQL = "insert into dynwrapx ("
		strSQL = strSQL & strCols 
		strSQL = strSQL & ") values ("
		strSQL = strSQL & strVals 
		strSQL = strSQL & ");"
		
		DBUtil_UpdateByNativeSQL g_objDBHelper, strSQL
	End Function 
	
	Public Function UpdateByMethod(strMethod, objDict)
		Dim strDst
		Dim strSQL
		
		strDst = DBUtil_GenerateStringByDict(objDict)
		
		strSQL = "update dynwrapx "
		strSQL = strSQL & "set "
		strSQL = strSQL & strDst 
		strSQL = strSQL & " where method='"
		strSQL = strSQL & strMethod
		strSQL = strSQL & "';"
		
		DBUtil_UpdateByNativeSQL g_objDBHelper, strSQL
	End Function 
	
	Public Function DeleteByMethod(strMethod)
		Dim strSQL
		
		strSQL = "delete from dynwrapx where method='"
		strSQL = strSQL & strMethod
		strSQL = strSQL & "';"
		
		DBUtil_UpdateByNativeSQL g_objDBHelper, strSQL
	End Function 
End Class

Class CommandService
	Public strDebugTag
	
	Private Sub Class_Initialize()
		strDebugTag = "CommandService："
	End Sub 
	
	Private Sub Class_Terminate()

	End Sub
	
	Public Function QueryByKeyword(strKeyword)
		Dim strSQL
		
		strSQL = "select * from command "
		strSQL = strSQL & "where keyword='"
		strSQL = strSQL & strKeyword 
		strSQL = strSQL & "';"
		
		Set QueryByKeyword = DBUtil_QueryByNativeSQL(g_objDBHelper, strSQL)
	End Function
	
	Public Function Insert(objDict)
		Dim strCols
		Dim strVals
		Dim strSQL
		
		strCols = DBUtil_GenerateStringByArray(objDict.Keys)
		strVals = DBUtil_GenerateStringByArray(objDict.Items)
		
		strSQL = "insert into command ("
		strSQL = strSQL & strCols 
		strSQL = strSQL & ") values ("
		strSQL = strSQL & strVals 
		strSQL = strSQL & ");"
		
		DBUtil_UpdateByNativeSQL g_objDBHelper, strSQL
	End Function 
	
	Public Function UpdateByKeyword(strKeyword, objDict)
		Dim strDst
		Dim strSQL
		
		strDst = DBUtil_GenerateStringByDict(objDict)
		
		strSQL = "update command "
		strSQL = strSQL & "set "
		strSQL = strSQL & strDst 
		strSQL = strSQL & " where keyword='"
		strSQL = strSQL & strKeyword
		strSQL = strSQL & "';"
		
		DBUtil_UpdateByNativeSQL g_objDBHelper, strSQL
	End Function 
	
	Public Function DeleteByKeyword(strKeyword)
		Dim strSQL
		
		strSQL = "delete from command where keyword='"
		strSQL = strSQL & strKeyword
		strSQL = strSQL & "';"
		
		DBUtil_UpdateByNativeSQL g_objDBHelper, strSQL
	End Function 
End Class 

Class MusicService
	Public strDebugTag
	
	Private Sub Class_Initialize()
		strDebugTag = "MusicService："
	End Sub 
	
	Private Sub Class_Terminate()

	End Sub
	
	Public Function Query()
		Dim strSQL
		
		strSQL = "select * from music"
		
		Set Query = DBUtil_QueryByNativeSQL(g_objDBHelper, strSQL)
	End Function

	Public Function Insert(objDict)
		Dim strCols
		Dim strVals
		Dim strSQL
		
		strCols = DBUtil_GenerateStringByArray(objDict.Keys)
		strVals = DBUtil_GenerateStringByArray(objDict.Items)
		
		strSQL = "insert into music ("
		strSQL = strSQL & strCols 
		strSQL = strSQL & ") values ("
		strSQL = strSQL & strVals 
		strSQL = strSQL & ");"
		
		DBUtil_UpdateByNativeSQL g_objDBHelper, strSQL
	End Function
	
	Public Function UpdateByID(strID, objDict)
		Dim strDst
		Dim strSQL
		
		strDst = DBUtil_GenerateStringByDict(objDict)
		
		strSQL = "update music "
		strSQL = strSQL & "set "
		strSQL = strSQL & strDst 
		strSQL = strSQL & " where id='"
		strSQL = strSQL & strID
		strSQL = strSQL & "';"
		
		DBUtil_UpdateByNativeSQL g_objDBHelper, strSQL
	End Function 
	
	Public Function DeleteByID(strID)
		Dim strSQL
		
		strSQL = "delete from music where id='"
		strSQL = strSQL & strID
		strSQL = strSQL & "';"
		
		DBUtil_UpdateByNativeSQL g_objDBHelper, strSQL
	End Function 
End Class 

Class DBHelper
	Public strDebugTag
	Public strDataBase
	Public strConn
	Public objConn

	Private Sub Class_Initialize()
		strDebugTag = "DBHelper："
	End Sub 
	
	Private Sub Class_Terminate()
		DisConnect
	End Sub
	
	Public Sub Connect(strDataBase)
		On Error Resume Next 
        Err.Clear
        
        Me.strDataBase = strDataBase
		strConn = "DRIVER={SQLite3 ODBC Driver};"
		strConn = strConn & "DataBase=" & strDataBase & ";"
		
		Set objConn = CreateObject("ADODB.Connection")
		objConn.Open strConn
		
		If Err Or objConn.State = 0 Then 
			Debug.WriteLine strDebugTag, "数据库 ", strDataBase, " 连接失败！"
			Set objConn = Nothing
			Exit Sub 
		End If 
		
		Debug.WriteLine strDebugTag, "数据库 ", strDataBase, " 连接成功！"
	End Sub 
	
	Public Sub DisConnect() 
		If objConn = Empty Then  
            Exit Sub  
        End If
        If objConn Is Nothing Then 
        	Exit Sub
        End If 
        
        objConn.Close
        Set objConn = Nothing
        
        Debug.WriteLine strDebugTag, "数据库 ", strDataBase, " 释放成功！"
	End Sub 
	
	Public Function Query(strSQL)
		On Error Resume Next 
        Err.Clear
        
		Dim objCommand
		Dim objRecordSet
		Dim objList
		Dim objDict
		Dim objTmp
		
	    Set objCommand = CreateObject("ADODB.COMMAND")
	    objCommand.CommandText = strSQL
	    objCommand.ActiveConnection = objConn
	    Set objRecordSet = CreateObject("ADODB.RECORDSET")
		objRecordSet.Open objCommand
		
		Set objList = Collections_NewArrayList()
		
		If Err Then 
			Debug.WriteLine strDebugTag, "查询操作 ", strSQL, _
				" 失败。错误：", Err.Description
			Set Query = objList
			Exit Function
		End If 
		
		Do Until objRecordSet.EOF
			Set objDict = Dict_NewDict()
			For Each objTmp In objRecordSet.Fields
				objDict.Add objTmp.Name, objTmp.Value
			Next
			objList.Add objDict
			objRecordSet.MoveNext
		Loop
		
		objRecordSet.Close
		
		Set objRecordSet = Nothing 
		Set objTmp = Nothing 
		Set Query = objList
		
		Debug.WriteLine strDebugTag, "查询操作 ", strSQL, _
			" 成功，返回", objList.Count, "条数据！"
	End Function 
	
	Public Function Update(strSQL)
		On Error Resume Next 
        Err.Clear
        
		Dim objCommand
		Dim objRecordSet
        Dim objField
        
        Set objCommand = CreateObject("ADODB.COMMAND") 
        objCommand.CommandText = strSQL 
        objCommand.ActiveConnection = objConn 
        Set objRecordSet = CreateObject("ADODB.RECORDSET") 
        Set objRecordSet = objCommand.Execute
        
        If Err Then 
			Debug.WriteLine strDebugTag, "更新操作 ", strSQL, _
				" 失败。错误：", Err.Description
			Exit Function
		End If 
		
        Set objCommand = Nothing 
        Set objRecordSet = Nothing 
        
        Debug.WriteLine strDebugTag, "更新操作 ", strSQL, _
				" 成功！", Err.Description
	End Function 
End Class 

Public Function DBUtil_Filter(strInput)
	If IsNull(strInput) Then 
		Exit Function 
	End If 
	
	strInput = Replace(strInput, "'", "''")
	strInput = Replace(strInput, "/", "//")
	strInput = Replace(strInput, "[", "/[")
	strInput = Replace(strInput, "]", "/]")
	strInput = Replace(strInput, "%", "/%")
	strInput = Replace(strInput, "&", "/&")
	strInput = Replace(strInput, "_", "/_")
	strInput = Replace(strInput, "(", "/(")
	strInput = Replace(strInput, ")", "/)")
	
	DBUtil_Filter = strInput
End Function 

Public Function DBUtil_Connect(objDBHelper, strDBFile)
	On Error Resume Next 
	
	objDBHelper.Connect(strDBFile)
	If Err Then 
		MsgBox "DBUtil_Connect：" & Err.Description
		WScript.Quit
	End If
End Function 

Public Function DBUtil_QueryByNativeSQL(objDBHelper, strSQL)
	On Error Resume Next

	Dim objList
	
	Set objList = objDBHelper.Query(strSQL)
	If Err Then 
		MsgBox "DBUtil_QueryByNativeSQL：" & Err.Description
		WScript.Quit
	End If
	
	Set DBUtil_QueryByNativeSQL = objList
	Set objList = Nothing 
End Function 

Public Function DBUtil_UpdateByNativeSQL(objDBHelper, strSQL)
	On Error Resume Next

	objDBHelper.Update(strSQL)
	If Err Then 
		MsgBox "DBUtil_UpdateByNativeSQL：" & Err.Description
		WScript.Quit
	End If
End Function 
	
Public Function DBUtil_GenerateStringByArray(arrData)
	Dim strDst
	Dim i
	
	For i = 0 To UBound(arrData)
		If i <> 0 Then 
			strDst = strDst & ", "
		End If
		strDst = strDst & "'" & arrData(i) & "'"
	Next
	
	DBUtil_GenerateStringByArray = strDst
End Function 

Public Function DBUtil_GenerateStringByDict(objDict)
	Dim strDst
	Dim i
	
	For i = 0 To objDict.Count - 1
		If i <> 0 Then 
			strDst = strDst & ", "
		End If 
		strDst = strDst & objDict.Keys()(i)
		strDst = strDst & "='"
		strDst = strDst & objDict.Items()(i)
		strDst = strDst & "'"
	Next
	
	DBUtil_GenerateStringByDict = strDst 
End Function

Function Clipboard_GetData()
	On Error Resume Next 
	
	Dim strText
	
	strText = g_objHtml.ParentWindow.ClipboardData.GetData("Text")
	If IsNull(strText) Then
		Debug.WriteLine "Clipboard_GetData：", "获取剪贴板数据出错，请重试"
		strText = ""
	End If 
	
	Clipboard_GetData = strText
End Function 

Function Math_Rnd(intMin, intMax)  
	Randomize
	Math_Rnd = Int((intMax - intMin + 1)*Rnd() + intMin)
End Function

Function Dict_NewDict()
	Set Dict_NewDict = CreateObject("Scripting.Dictionary")
End Function

Function Array_IndexOf(arrData, varData, blnFirst)
	Dim varTemp
	Dim intIndex
	Dim i
	
	intIndex = -1
	For i = 0 To UBound(arrData)
		If IsObject(varData) Then 
			Set varTemp = arrData(i)
			If varTemp Is varData Then 
				intIndex = i
			End If 
		Else 
			varTemp = arrData(i)
			If varTemp = varData Then 
				intIndex = i
			End If 
		End If 
		If blnFirst And intIndex <> -1 Then 
			Exit For 
		End If 
	Next
	
	Array_IndexOf = intIndex
End Function

Function Array_FirstIndexOf(arrData, varData)
	Dim intIndex
	
	intIndex = Array_IndexOf(arrData, varData, True)
	
	Array_FirstIndexOf = intIndex
End Function

Function Array_LastIndexOf(arrData, varData)
	Dim intIndex
	
	intIndex = Array_IndexOf(arrData, varData, False)
	
	Array_LastIndexOf = intIndex
End Function

Function Array_ToList(arrData)
	Dim objList
	Dim varItem
	
	Set objList = Collections_NewArrayList()
	For Each varItem In arrData
	    objList.Add varItem
	Next
	
	Set Array_ToList = objList
	Set objList = Nothing
End Function

Function Array_ToArray(objList)
	Dim arrData
	Dim intBound
	
	intBound = objList.Count - 1
	ReDim Preserve arrData(intBound)
	
	For i = 0 To intBound
	    arrData(i) = objList(i)
	Next
	
	Array_ToArray = arrData
End Function

Function Collections_NewArrayList()

	Set Collections_NewArrayList = CreateObject("System.Collections.ArrayList")
	
End Function 

Function Collections_NewQueue()

	Set Collections_NewQueue =  CreateObject("System.Collections.Queue")
	
End Function 

Function Collections_NewStack()

	Set Collections_NewStack =  CreateObject("System.Collections.Stack")
	
End Function 

Function File_LoadFile(strFile, bytBuff)
	Dim objStream
	
	If Not g_objFSO.FileExists(strFile) Then
		File_LoadFile = False
		Exit Function 
	End If 
	
	Set objStream = g_objFSO.OpenTextFile(strFile, ForReading)  
	bytBuff = objStream.ReadAll
	File_LoadFile = True
	objStream.Close
	
	Set objStream = Nothing
End Function

Function File_MD5(strFile)
	Dim bytBuff
	Dim blnSucc
	Dim strMD5
	
	blnSucc = File_LoadFile(strFile, bytBuff)
	If Not blnSucc Then 
		Exit Function
	End If 
	
	strMD5 = CAPICOM_HashedData(bytBuff, CAPICOM_HASH_ALGORITHM_MD5)
	File_MD5 = strMD5
End Function

Function CAPICOM_HashedData(bytContent, intAlgorithm)
	Dim objHashedData
	Dim strHash
	
	Set objHashedData = CreateObject("CAPICOM.HashedData")
	objHashedData.Algorithm = intAlgorithm
	objHashedData.Hash bytContent
	strHash = objHashedData.Value
	CAPICOM_HashedData = LCase(strHash)
	
	Set objHashedData = Nothing
End Function

Function HTTP_DoHttp(strMethod, strURL, varBody, varAsync)
	Dim objHttp
	
	Set objHttp = CreateObject("Msxml2.XMLHTTP")
	objHttp.open strMethod, strURL, varAsync
	objHttp.send varBody
	HTTP_DoHttp = objHttp.responseText
	
	Set objHttp = Nothing
End Function

Function HTTP_Get(strURL, varBody)
	HTTP_Get = HTTP_DoHttp("GET", strURL, varBody, False)
End Function 

Function HTTP_Post(strURL, varBody)
	HTTP_Post = HTTP_DoHttp("POST", strURL, varBody, False)
End Function

Function HTTP_Tuling(strKey, strInfo, strLoc, strUserId)
	Dim strBody
	Dim strResp
	Dim objSC
	Dim objDict
	Dim objList
	Dim objTemp
	Dim arrKeys
	Dim intLen, i, j
	
	strBody = "{'key':'#KEY#', 'info':'#INFO#', 'loc':'#LOC#', 'userid':'#USERID#'}"
	strBody = Replace(strBody, "'", """")
	strBody = Replace(strBody, "#KEY#", strKey)
	strBody = Replace(strBody, "#INFO#", strInfo)
	strBody = Replace(strBody, "#LOC#", strLoc)
	strBody = Replace(strBody, "#USERID#", strUserId)
	strResp = HTTP_Post(STR_TULING_API_CONF_URL, strBody)
	
	Set objSC = CreateObject("MSScriptControl.ScriptControl")
	objSC.Language = "JScript"
	objSC.AddCode "var ret = " & strResp & ";"
    
   	Set objDict = Dict_NewDict
   	objDict.Add "code", objSC.Eval("ret.code")
	objDict.Add "text", objSC.Eval("ret.text")
	objDict.Add "url" , objSC.Eval("ret.url")
   	
   	If objDict("code") = TULING_API_CODE_NEWS Then 
   		arrKeys = Array("article", "source", "icon", "detailurl")
   	ElseIf objDict("code") = TULING_API_CODE_COOK Then 
   		arrKeys = Array("name", "icon", "info", "detailurl")
   	End If 
   	
   	If Not IsEmpty(arrKeys) Then 
    	Set objList = Collections_NewArrayList
		intLen = objSC.Eval("ret.list.length")
		For i = 0 To intLen -1
			Set objTemp = Dict_NewDict
			For j = 0 To UBound(arrKeys)
				objTemp.Add arrKeys(j), objSC.Eval("ret.list[" & i &"]." & arrKeys(j))
			Next 
			objList.Add objTemp
		Next
		objDict.Add "list", objList
   	End If 

   	Set HTTP_Tuling = objDict
   	
	Set objSC = Nothing
	Set objDict = Nothing
	Set objList = Nothing
	Set objTemp = Nothing
End Function 

Function HTTP_Tuling_Ask(strInfo)
	Set HTTP_Tuling_Ask = HTTP_Tuling(STR_TULING_API_CONF_KEY, strInfo, _
		STR_TULING_API_CONF_LOC, STR_TULING_API_CONF_UID)
End Function 

Function SP_VOICE_Speak(strContent)
	On Error Resume Next 
	Dim objSpVoice

	Set objSpVoice = CreateObject("SAPI.SpVoice")
	Set objSpVoice.Voice = objSpVoice.GetVoices("Name=VW Hui").Item(0)
	objSpVoice.Speak strContent, 1 Or 2
	
	Set objSpVoice = Nothing
	On Error Goto 0
End Function

Function PropertyExists(objContext, strName)
	Dim blnExists
	
	On Error Resume Next
	Eval "objContext." & strName
	blnExists = Not CBool(Err)
	On Error Goto 0
	
	PropertyExists = blnExists
End Function
