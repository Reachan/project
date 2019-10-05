#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;-----------------------------------
;無変換キーを修飾キーに
;-----------------------------------
vk1D & E::Send, {Blind}{UP}
vk1D & D::Send, {Blind}{DOWN}
vk1D & S::Send, {Blind}{LEFT}
vk1D & F::Send, {Blind}{RIGHT}

vk1D & `;::Send, {Blind}{Backspace} ;無変換 + ; = Backspace
vk1D & M::Send, {Blind}{Delete}     ;無変換 + M = Delete
vk1D & O::Send, {Blind}{PgUp}       ;無変換 + O = PgUp
vk1D & L::Send, {Blind}{PgDn}       ;無変換 + L = PgDn
vk1D & I::Send, {Blind}{Home}       ;無変換 + I = Home
vk1D & K::Send, {Blind}{End}        ;無変換 + K = End
vk1D & vkF3::Send, {Blind}{Esc}     ;無変換 + 半角/全角 = Ese※IMEのON/OFFで発生するイベントが違うため、二つとも定義
vk1D & vkF4::Send, {Blind}{Esc}     ;無変換 + 半角/全角 = Ese※IMEのON/OFFで発生するイベントが違うため、二つとも定義

vk1D & 1::Send, {Blind}{F1}      ;無変換 + 1 = F1
vk1D & 2::Send, {Blind}{F2}      ;無変換 + 2 = F2
vk1D & 3::Send, {Blind}{F3}      ;無変換 + 3 = F3
vk1D & 4::Send, {Blind}{F4}      ;無変換 + 4 = F4
vk1D & 5::Send, {Blind}{F5}      ;無変換 + 5 = F5
vk1D & 6::Send, {Blind}{F6}      ;無変換 + 6 = F6
vk1D & 7::Send, {Blind}{F7}      ;無変換 + 7 = F7
vk1D & 8::Send, {Blind}{F8}      ;無変換 + 8 = F8
vk1D & 9::Send, {Blind}{F9}      ;無変換 + 9 = F9
vk1D & 0::Send, {Blind}{F10}     ;無変換 + 0 = F10
vk1D & -::Send, {Blind}{F11}     ;無変換 + - = F11
vk1D & ^::Send, {Blind}{F12}     ;無変換 + ^ = F12
vk1D & vk1C:: AppsKey			 ;無変換 + 変換 = AppsKey

;-----------------------------------
;変換キーを修飾キーに
;-----------------------------------
;vk1C & I::MouseMove, 0, -50, 0, R
vk1C & J::MouseMove, -50, 0, 0, R
vk1C & K::MouseMove, 0, 50, 0, R
vk1C & L::MouseMove, 50, 0, 0, R

vk1C & I::	
	KeyWait,I,T0.2　;0.2秒対象キーが押されるのを待つ
	If(ErrorLevel)
	{	
		;1度押し
		MouseMove, 0, -10, 0, R
	}
	;2度押し	
	MouseMove, 0, -50, 0, R
	Return


;　書き方をメモするために残す
;vk1D & D::
;	If GetKeyState("Shift", "P"){
;	    Send, +{DOWN}
;	}else{
;		Send {DOWN}
;	}
;	Return


;************************************
;-----------------------------------
;共通スクリプト
;-----------------------------------
;************************************

;-----------------------------------
;選択したファイルのバックアップを作成
;バックアップファイル名にファイルの最終編集日付をつける
;「変換キー」＋D
;-----------------------------------
vk1C & D::
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
	ClipWait, 2
}
return

;-----------------------------------
;「変換キー」+nで名前入力
;一回目は漢字、二回目は片仮名、長押しはローマ字
;-----------------------------------
vk1C & n::
	KeyWait,n,T0.3　;0.3秒対象キーが押されたかどうか
	If(ErrorLevel)
	{	;長押し
		Send,aiueo
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

;-----------------------------------
;すべての列のサイズを自動的に変更する
;普通は「ctrl」と「+」を同時に押せばできる
;-----------------------------------
vk1C & /::
	Send ^{NumpadAdd}
return


;************************************
;-----------------------------------
;各アプリケーション用のスクリプト
;-----------------------------------
;************************************

;-----------------------------------
;Alt+1で書式なしペースト（Excel以外に適用）
;-----------------------------------
#IfWinNotActive ("ahk_exe EXCEL.EXE")
!1::
	ClipSaved := ClipboardAll ;save original clipboard contents
	clipboard = %clipboard% ;remove formatting
	Send ^v ;send the Ctrl+V command
	Clipboard := ClipSaved ;restore the original clipboard contents
	ClipSaved = ;clear the variable
	Return
#IfWinNotActive

;-----------------------------------
;コマンドプロンプト用
;-----------------------------------
;ctrl+vで張り付け
#If WinActive("ahk_class ConsoleWindowClass")
^v::Send,!{Space}ep	;貼り付け
	return
#IfWinActive

;-----------------------------------
;オブジェクトブラウザ用(使えない)
;-----------------------------------
;「変換キー」+pgUPで画面変換
#If WinActive("ahk_exe ob13.exe")
^PgUp::Send +^{Tab}
	return
^PgDn::Send ^{Tab}
	return
#IfWinActive

;-----------------------------------
;さくらエディター用
;-----------------------------------
;Ctrl+Pgup/PgDnでタグ切り替え
;#If WinActive("ahk_class TextEditorWindowWP172")
;	^PgUp::Send +^{Tab}
;		return
;	^PgDn::Send ^{Tab}
;		return
;#IfWinActive

;-----------------------------------
;Chrome用
;-----------------------------------
;Ctrl+Pgup/PgDnでタグ切り替え
#If WinActive("ahk_exe chrome.exe")
^PgUp::Send +^{Tab}
	return
^PgDn::Send ^{Tab}
	return
#IfWinActive

;-----------------------------------
;Excel用
;-----------------------------------
#If WinActive("ahk_exe EXCEL.EXE")
	;文字の色を赤にする
	;vk1C & R::Send {Alt}HFC{Down}{Down}{Down}{Down}{Down}{Down}{Down}{LEFT}{LEFT}{LEFT}{LEFT}{ENTER}
	;	return
	;「変換キー」+rでセル書式変更（一回目は赤字、二回目は黒字、長押しはセルを黄色）
	vk1C & r::
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


	;「変換キー」+wで画面拡大(一回目は130%、二回目は150%、長押しは100%）
	vk1C & w::
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
	vk1C & Y::Send {Alt}Y2Y7{Down}{Enter}
		return
	
	;選択した対象をグループ化する
	vk1C & G::Send {AppsKey}GG
		return
	
	;選択した枠内の境界線 (横) を削除または適用する。
	vk1C & 1::
		Send ^1
		Send !H
		Send {Enter}
		return	
	
	;フィルタをクリア
	vk1C & 2::
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

;-----------------------------------
;PPT用
;-----------------------------------
#If WinActive("ahk_exe POWERPNT.EXE")
	;文字の色を赤にする
	vk1C & R::Send {Alt}HFC{Down}{Down}{Down}{Down}{Down}{Down}{RIGHT}{ENTER}
		return
#If WinActive

