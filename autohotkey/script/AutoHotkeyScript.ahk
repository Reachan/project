#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Alt+1で書式なしペースト（Excel以外に適用）
#IfWinNotActive ("ahk_exe EXCEL.EXE")
!1::
	ClipSaved := ClipboardAll ;save original clipboard contents
	clipboard = %clipboard% ;remove formatting
	Send ^v ;send the Ctrl+V command
	Clipboard := ClipSaved ;restore the original clipboard contents
	ClipSaved = ;clear the variable
	Return
#IfWinNotActive

;選択したファイルのバックアップを作成
;バックアップファイル名にファイルの最終編集日付をつける
;無変換＋D
vkF2sc070 & D::
{
	;clipboardの内容を退避する
	temp = %clipboard%

	;選択したファイルをコピーする。windows上はファイルをコピーする時に、コピーしたのはファイルのパスのため、この方法でファイルパスを取得
	clipboard = ;初期化
	Send, ^c
	ClipWait ;クリップボードの内容がテキストとして読み取れるものになるのを待つ。
	;コピーしたファイルパスを変数に格納する
	SelectedFile = %clipboard% 
	;MsgBox,%SelectedFile%
	
	;ファイルパスで、ファイル名、パス、拡張子、拡張子なしのファイル名、ドライブ名（CドライブかDドライブか）の情報を取得
	SplitPath, SelectedFile, name, dir, ext, name_no_ext, drive
	
	;ファイルの最終更新時間を取得
	FileGetTime, LastModify,%SelectedFile%,M
	;最終更新時間のフォーマットを編集
	FormatTime, LastModify, %LastModify%, _yyyyMMdd_HHmmss
	
	;上記処理で取得した情報を基にバックアップファイルを作成
	FileCopy, %SelectedFile%, %dir%\%name_no_ext%%LastModify%.%ext%
	
	;退避したclipboardの内容を復元
	Clipboard  := temp
}
return

;vkF2sc070 & S::
;send ^c
;	SelectedFile = %clipboard%
;	SplitPath, %clipboard%, Name, Dir, Ext,Name_No_Ext, Drive
;	MsgBox, %Name%
;Return

;マウスカーソル移動（fast）：無変換＋Up,Down,Left,Right
;vkF2sc070 & Up:: 
;	MouseMove 0,-50,0,R
;return
;vkF2sc070 & Down:: 
;	MouseMove 0,50,0,R
;return
;vkF2sc070 & Left:: 
;	MouseMove -50,0,0,R
;return
;vkF2sc070 & Right:: 
;	MouseMove 50,0,0,R
;return

;マウスカーソル移動(slow):RCtrl+Up,Down,Left,Right　
;RCtrl & Up:: 
;	MouseMove 0,-10,0,R
;return
;RCtrl & Down:: 
;	MouseMove 0,10,0,R
;return
;RCtrl & Left:: 
;	MouseMove -10,0,0,R
;return
;RCtrl & Right:: 
;	MouseMove 10,0,0,R
;return

;方向キー　：無変換＋E,D,S,F
;vkF2sc070 & E:: Send {UP}
;		return
;vkF2sc070 & D:: Send {DOWN}
;		return
;vkF2sc070 & S:: Send {LEFT}
;		return
;vkF2sc070 & F:: Send {RIGHT}
;		return

;「をペアで
vkF2sc070 & [::Send {[}{]}{ENTER}{LEFT}
		return


;AppsKey
vk1Csc079:: AppsKey

;CapsLockキーにCtrlキーの仕事をさせる
;sc03a::Ctrl

;すべての列のサイズを自動的に変更する
;普通は「ctrl」と「+」を同時に押せばできる
vkF2sc070 & /::
	Send ^{NumpadAdd}
return

;左クリック
;vkF2sc070 & Shift::
;	MouseClick,left
;return

;右クリック
;vkF2sc070 & RCtrl::
;	MouseClick,right
;return

;コマンドプロンプト用
;ctrl+vで張り付け
#If WinActive("ahk_class ConsoleWindowClass")
^v::Send,!{Space}ep	;貼り付け
	return
#IfWinActive

;オブジェクトブラウザ用(使えない)
;無変換+pgUPで画面変換
#If WinActive("ahk_exe ob13.exe")
^PgUp::Send +^{Tab}
	return
^PgDn::Send ^{Tab}
	return
#IfWinActive

;ONENOTE用
;Ctrl+Pgup/PgDnでタグ切り替え
#If WinActive("ahk_class Framework::CFrame")
^PgUp::Send +^{Tab}
	return
^PgDn::Send ^{Tab}
	return
;文字の色を赤にする
vkF2sc070 & R::Send {Alt}HFC{Down}{Down}{Down}{Down}{Down}{Down}{Down}{LEFT}{LEFT}{LEFT}{LEFT}{ENTER}
	return
#IfWinActive

;VisualBasic 2013用
;Ctrl+Pgup/PgDnでタグ切り替え
#If WinActive("ahk_exe devenv.exe")
^PgUp::Send +^{Tab}
	return
^PgDn::Send ^{Tab}
	return
#IfWinActive

;Ctrl+;で日付入力(一回目はyyyy/mm/dd形式、二回目はyyyymmdd形式、長押しはyyyyMMdd_HHmmss_形式）
^vkBBsc027::
	KeyWait,vkBBsc027,T0.3　;0.3秒対象キーが押されたかどうか
	If(ErrorLevel)
	{	;長押し
		FormatTime,TimeString,,yyyyMMdd_HHmmss
		Send,%TimeString%_
		KeyWait,vkBBsc027
		Return
	}	
	KeyWait,vkBBsc027,D T0.2　;0.2秒対象キーが押されるのを待つ
	If(ErrorLevel)
	{	
		;1度押し
		FormatTime,TimeString,,yyyy/MM/dd
		Send,%TimeString%
		KeyWait,vkBBsc027
		Return
	}
	;2度押し	
	FormatTime,TimeString,,yyyyMMdd
	Send,%TimeString%
	KeyWait,vkBBsc027
	Return

;さくらエディター用
;Ctrl+Pgup/PgDnでタグ切り替え
;20190618 start
;さくらエディタで設定可能のため、いったんコメントアウト。
;#If WinActive("ahk_class TextEditorWindowWP172")
;	^PgUp::Send +^{Tab}
;		return
;	^PgDn::Send ^{Tab}
;		return
#IfWinActive
;20190618 end

;無変換+nで名前入力(一回目は漢字、二回目は片仮名、長押しはローマ字）
vkF2sc070 & n::
	KeyWait,n,T0.3　;0.3秒対象キーが押されたかどうか
	If(ErrorLevel)
	{	;長押し
		Send,XXXX
		KeyWait,n
		Return
	}	
	KeyWait,n,D T0.2　;0.2秒対象キーが押されるのを待つ
	If(ErrorLevel)
	{	
		;1度押し
		Send,あいうえお
		KeyWait,n
		Return
	}
	;2度押し	
	Send,アイウエオ
	KeyWait,n
	Return



;Chrome用
;Ctrl+Pgup/PgDnでタグ切り替え
#If WinActive("ahk_exe chrome.exe")
^PgUp::Send +^{Tab}
	return
^PgDn::Send ^{Tab}
	return
#IfWinActive

;Excel用
#If WinActive("ahk_exe EXCEL.EXE")
	;文字の色を赤にする
	;vkF2sc070 & R::Send {Alt}HFC{Down}{Down}{Down}{Down}{Down}{Down}{Down}{LEFT}{LEFT}{LEFT}{LEFT}{ENTER}
	;	return
	;無変換+rでセル書式変更（一回目は赤字、二回目は黒字、長押しはセルを黄色）
	vkF2sc070 & r::
		KeyWait,r,T0.3　;0.3秒対象キーが押されたかどうか
		If(ErrorLevel)
		{	;長押し
			Send, {Alt}HH{Down}{Down}{Down}{Down}{Down}{Down}{Right}{Right}{Right}{ENTER}
			KeyWait,r
			Return
		}	
		KeyWait,r,D T0.1　;0.2秒対象キーが押されるのを待つ
		If(ErrorLevel)
		{	
			;1度押し
			Send, {Alt}HFC{Down}{Down}{Down}{Down}{Down}{Down}{Down}{LEFT}{LEFT}{LEFT}{LEFT}{ENTER}
			KeyWait,r
			Return
		}
		;2度押し	
		Send, {Alt}HFC{Enter}
		KeyWait,r
		Return


	;無変換+wで画面拡大(一回目は130%、二回目は150%、長押しは100%）
	vkF2sc070 & w::
		KeyWait,w,T0.3　;0.3秒対象キーが押されたかどうか
		If(ErrorLevel)
	{	;長押し
		Send,{Alt}wqc100{Enter}
		KeyWait,w
		Return
	}	
	KeyWait,w,D T0.2　;0.2秒対象キーが押されるのを待つ
	If(ErrorLevel)
	{	
		;1度押し
		Send,{Alt}wqc130{Enter}
		KeyWait,w
		Return
	}
	;2度押し	
	Send,{Alt}wqc150{Enter}
	KeyWait,w
	Return
	
	;すべてのシールのコンソールをA1にする
	vkF2sc070 & Y::Send {Alt}Y2Y7{Down}{Enter}
		return
	
	;選択した対象をグループ化する
	vkF2sc070 & G::Send {AppsKey}GG
		return
	
	;選択した枠内の境界線 (横) を削除または適用する。
	vkF2sc070 & 1::
		Send ^1
		Send !H
		Send {Enter}
		return	
	
	;フィルタをクリア
	vkF2sc070 & 2::
		Send {Alt}HSC
		return
			
	;書式のみ貼り付け
	;使えない
	;+^v::
		;形式を選択して貼り付けウィンドウーを開く
	;	Send ^!V
	;	Send !T
	;	Send {Enter}
	;	return
	
#If WinActive

;PPT用
#If WinActive("ahk_exe POWERPNT.EXE")
	;文字の色を赤にする
	vkF2sc070 & R::Send {Alt}HFC{Down}{Down}{Down}{Down}{Down}{Down}{RIGHT}{ENTER}
		return
#If WinActive

;ONENOTE用
;Ctrl+Pgup/PgDnでタグ切り替え
#If WinActive("ahk_class Framework::CFrame")
^PgUp::Send +^{Tab}
	return
^PgDn::Send ^{Tab}
	return
;文字の色を赤にする
vkF2 & R::Send {Alt}HFC{Down}{Down}{Down}{Down}{Down}{Down}{Down}{LEFT}{LEFT}{LEFT}{LEFT}{ENTER}
	return
#IfWinActive

;VisualBasic 2013用
;Ctrl+Pgup/PgDnでタグ切り替え
#If WinActive("ahk_exe devenv.exe")
^PgUp::Send +^{Tab}
	return
^PgDn::Send ^{Tab}
	return
#IfWinActive

;Ctrl+;で日付入力(一回目はyyyy/mm/dd形式、二回目はyyyymmdd形式、長押しはyyyyMMdd_HHmmss_形式）
^vkBB::
	KeyWait,vkBB,T0.3　;0.3秒対象キーが押されたかどうか
	If(ErrorLevel)
	{	;長押し
		FormatTime,TimeString,,yyyyMMdd_HHmmss
		Send,%TimeString%_
		KeyWait,vkBB
		Return
	}	
	KeyWait,vkBB,D T0.2　;0.2秒対象キーが押されるのを待つ
	If(ErrorLevel)
	{	
		;1度押し
		FormatTime,TimeString,,yyyy/MM/dd
		Send,%TimeString%
		KeyWait,vkBB
		Return
	}
	;2度押し	
	FormatTime,TimeString,,yyyyMMdd
	Send,%TimeString%
	KeyWait,vkBB
	Return
