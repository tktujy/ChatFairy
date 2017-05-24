Option Explicit

Const MAIN_LOOP_DELAY = 10

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

Dim g_strMode : g_strMode = STR_WORKING_MODE_TULING
Dim g_strDBPath
Dim g_lngTimestamp

Set g_objHtml= CreateObject("htmlfile")
Set g_objWshShell = CreateObject("WScript.Shell")
Set g_objFSO = CreateObject("Scripting.FileSystemObject")
Set g_objDBHelper = New DBHelper

g_strDBPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "data.db")

Call VBSMain()

Set g_objHtml = Nothing
Set g_objWshShell = Nothing 
Set g_objFSO = Nothing 
Set g_objDBHelper = Nothing

Function VBSMain
	Dim objIgnoreList
	Dim objList
	Dim objTemp
	Dim strNewData
	
	DBUtil_Connect g_objDBHelper, g_strDBPath
	
	Set objIgnoreList = Array_NewList()
	objIgnoreList.Add strNewData
	objIgnoreList.Add ","
	objIgnoreList.Add "，"
	objIgnoreList.Add ""
	objIgnoreList.Add "？"
	
	Set objList = Array_NewList()
	objList.Add New CommandPlugin       ' 命令模式
	objList.Add New MusicScannerPlugin  ' 音乐扫描
	objList.Add New MusicPlayerPlugin   ' 音乐模式
	objList.Add New TulingPlugin        ' 图灵机器人
	
	For Each objTemp In objList 
		objTemp.Plugin_Init
	Next
	
	Do
    	WScript.Sleep MAIN_LOOP_DELAY
    	g_lngTimestamp = g_lngTimestamp + MAIN_LOOP_DELAY
    	strNewData = Clipboard_GetData()
    	
    	If StrComp(strNewData, "退出程序") = 0 Then 
    		Exit Do 
    	End If
    	
    	If Not objIgnoreList.Contains(strNewData) Then
    		objIgnoreList.RemoveAt 0
    		objIgnoreList.Insert 0, strNewData
    		DispatchPlugin objList, strNewData
    	End If 
    	
    	If g_lngTimestamp Mod TIME_UNIT_S = 0 Then
	    	For Each objTemp In objList
	    		objTemp.Plugin_Timer g_strMode, strNewData
	    	Next 
		End If
    Loop

    Set objIgnoreList = Nothing 
    Set objList = Nothing
    Set objTemp = Nothing
End Function 

Function DispatchPlugin(objList, strKeyw)
	Dim objTemp
	Dim objModeKeys
	Dim blnHitMode
	Dim blnHitKeys
	
	For Each objTemp In objList
		Set objModeKeys = objTemp.ModeKeys
		
		blnHitMode = objModeKeys.HitMode(strKeyw) 
		blnHitKeys = objModeKeys.HitKeyw(strKeyw)
		
		If blnHitMode Then 
			Debug.WriteLine "模式切换  ，", g_strMode, " -> ", strKeyw
			g_strMode = strKeyw
		End If
		
		blnHitMode = objModeKeys.HitMode(g_strMode) 
		
		If blnHitMode And blnHitKeys Then 
			Debug.WriteLine "关键字命中，", strKeyw
			objTemp.Plugin_Handle g_strMode, strKeyw
		End If 
	Next
End Function

Class CommandPlugin
	Public strDebugTag
	Public objModeKeys
	Public objService
	
	Private Sub Class_Initialize()
		strDebugTag = "CommandPlugin："
		
		Set objModeKeys = New ModeKeysCls
		Set objService = New CommandService
		
		objModeKeys.RegMode STR_WORKING_MODE_COMMAND
		objModeKeys.RegKeywAll True
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
		
		If Len(strPerformArgs) > 0 Then 
			strCmdText = strCmdText & Space(1)
			strCmdText = strCmdText & strPerformArgs
		End If 
		
		If Len(strPerformObject) > 0 Then 
			strCmdText = strCmdText & Space(1)
			strCmdText = strCmdText & strPerformObject
		End If 
		
		Debug.WriteLine strDebugTag, "执行 ", strCmdText
		
		g_objWshShell.Run strCmdText
	End Function 
	
	Public Function Plugin_Timer(strMode, strKeyw)

	End Function 
End Class 

Class MusicScannerPlugin
	Public strDebugTag
	Public objModeKeys
	Public objService
	Public objDirList
	
	Private Sub Class_Initialize()
		strDebugTag = "MusicScannerPlugin："
		Set objModeKeys = New ModeKeysCls
		Set objService = New MusicService
		Set objDirList = Array_NewList
		
		objDirList.Add STR_PLUGIN_MUSIC_SCANNER_DIR
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
	
	Public Function Plugin_Handle(strMode, strKeyw)
	
	End Function 
	
	Public Function Plugin_Timer(strMode, strKeyw)
	
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
		Set objList = Array_NewList()
		
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
		
		objModeKeys.RegMode STR_WORKING_MODE_MUSIC
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_RANDOM
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_PAUSE
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_STOP
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_GO
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_NEXT
		objModeKeys.RegKeyw STR_PLUGIN_MUSIC_PLAYER_PREV
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
		Set objService = Nothing
		Set objPlayer = Nothing
	End Sub
	
	Public Property Get ModeKeys()
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function Plugin_Init()
		Dim objList
		Dim objTemp
		
		Debug.WriteLine strDebugTag, "Init"
		
		Set objList = objService.Query()
		For Each objTemp In objList 
			objDict.Add objTemp("id"), objTemp("path")
		Next 
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
		If Not objModeKeys.HitMode(strMode) Then 
			Exit Function
		End If 
		
		If Not objModeKeys.HitKeyw(Me.strKeyw) Then 
			Exit Function
		End If 
	
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
	Public objModeKeys
	
	Private Sub Class_Initialize()
		Dim arrKeyw
		
		strDebugTag = "TulingPlugin："
		Set objModeKeys = New ModeKeysCls
		objModeKeys.RegMode STR_WORKING_MODE_TULING
		objModeKeys.RegKeywAll True
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
	End Sub 
	
	Public Property Get ModeKeys()
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function Plugin_Init()
		Debug.WriteLine strDebugTag, "Init"
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
	
	Public Function Plugin_Timer(strMode, strKeyw)

	End Function 
End Class 

' 不要New，方便写新模块时使用
Class SamplesPlugin
	Public strDebugTag
	Public objModeKeys
	
	Private Sub Class_Initialize()
		Dim arrKeyw
		
		strDebugTag = "SamplesPlugin："
		Set objModeKeys = New ModeKeysCls
		objModeKeys.RegMode "例子模式"
		objModeKeys.RegKeywAll True
	End Sub 
	
	Private Sub Class_Terminate()
		Set objModeKeys = Nothing
	End Sub 
	
	Public Property Get ModeKeys()
		Set ModeKeys = objModeKeys
	End Property 
	
	Public Function Plugin_Init()
		Debug.WriteLine strDebugTag, "Init"
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
		Set objMode = Array_NewList()
		Set objKeys = Array_NewList()
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
		
		Set objList = Array_NewList()
		
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
	
    Set objList = Array_NewList()
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

Function Array_NewList()
	Set Array_NewList = CreateObject("System.Collections.ArrayList")
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
    	Set objList = Array_NewList
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

Sub SP_VOICE_Speak(strContent)
	Dim objSpVoice
	
	Set objSpVoice = CreateObject("SAPI.SpVoice") 
	objSpVoice.Speak strContent
	
	Set objSpVoice = Nothing
End Sub
