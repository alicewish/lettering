;这是同步用AHK脚本——墨问非名
;全局快捷键
RAlt & j::AltTab
RAlt & k::ShiftAltTab
^+#!B::run C:\Users\Alicewish\AppData\Roaming\baidu\BaiduYunGuanjia\baiduyunguanjia.exe
^+#!C::run C:\Program Files (x86)\Calibre2\calibre.exe
^+#!D::run C:\Program Files (x86)\DuplicateCleanerPro\DuplicateCleaner.exe
^+#!E::run C:\Program Files\Everything\Everything.exe
^+#!F::run C:\Program Files (x86)\Mozilla Firefox\firefox.exe
^+#!H::run \\Mac\Host\Volumes\Mack\汉化
^+#!I::run C:\Program Files (x86)\Internet Download Manager\IDMan.exe
^+#!L::run C:\Program Files (x86)\iTools\iTools.exe
^+#!N::run C:\Program Files (x86)\ABBYY FineReader 12\FineReader.exe
^+#!O::run C:\Program Files\totalcmd\TOTALCMD64.EXE
^+#!P::run C:\Program Files (x86)\Adobe\Photoshop CS6\Photoshop.exe
^+#!Q::run C:\Program Files (x86)\Tencent\QQ\Bin\QQScLauncher.exe
^+#!R::run C:\Program Files (x86)\Advanced Renamer\ARen.exe
^+#!T::run C:\Program Files (x86)\Tongbu\Tongbu.exe
^+#!U::run C:\Program Files\UltraEdit\Uedit32.exe
^+#!Y::run C:\Users\Alicewish\AppData\Local\Youdao\Dict\Application\YodaoDict.exe
^+#!z::
FormatTime, TimeString, yyyy/MM/dd hh:mm:ss tt R
MsgBox 目前的时间是%TimeString%。%A_ComputerName%%A_UserName%
{
FormatTime, TimeStringStart, yyyy/MM/dd hh:mm:ss tt R
IfWinExist, ahk_class OpusApp
Sleep 1000 ;延时1秒
WinActivate ; 使用前面找到的窗口
WinGetTitle, Title ;获取窗口名
Sleep 1000 ;延时1秒
FormatTime, TimeStringEnd, yyyy/MM/dd hh:mm:ss tt R
FileAppend,
(

填字项目：%Title%
开始时间：%TimeStringStart%
完成时间：%TimeStringEnd%

), \\Mac\Home\Documents\填字完成.txt
return
SoundBeep, 750, 500 ;以较高的音高进行发音并持续半秒.
Sleep 1000 ;延时1秒
/*
; 下面的函数把指定的秒数转换成相应的
; 小时数, 分钟数和秒数 (hh:mm:ss 格式).
MsgBox % FormatSeconds(7384)  ; 7384 = 2 小时 + 3 分钟 + 4 秒. 它的结果: 2:03:04
FormatSeconds(NumberOfSeconds)  ; 把指定的秒数转换成 hh:mm:ss 格式.
{
    time = 19990101  ; 任意日期的 *午夜*.
    time += %NumberOfSeconds%, seconds
    FormatTime, mmss, %time%, mm:ss
    return NumberOfSeconds//3600 ":" mmss
    ; 和上面方法不同的是，这里不支持超过 24 小时的秒数：
    ;FormatTime, hmmss, %time%, h:mm:ss
    ;return hmmss
}
SoundPlay *-1 ;简单的哗哗声. 如果声卡不可用, 则使用扬声器生成这个声音.
Sleep 1000 ;延时1秒
SoundPlay *16 ;手型（停止/错误声）
Sleep 1000 ;延时1秒
SoundPlay *32 ;问号声
Sleep 1000 ;延时1秒
SoundPlay *48 ;感叹声
Sleep 1000 ;延时1秒
SoundPlay *64 ;星号（消息声）
*/
}
Return
^+#!1::run \\Mac\Dropbox
^+#!2::run C:\Users\Alicewish\Google 云端硬盘
^+#!3::run \\Mac\Home\Downloads
^+#!4::run G:\Comics
^+#!5::run G:\熟肉
^+#!6::run \\Mac\Home\Documents
^+#!7::run \\Mac\Home\Downloads\Compressed
^+#!8::run \\Mac\Home\Documents\Tencent Files\782841300\FileRecv
^+#!9::run \\Mac\Home\Desktop
^+#!0:: ;激活MangaMeeya
{
IfWinExist, ahk_class MangaMeeya
    WinActivate ; 使用前面找到的窗口
Return
}
^+#!-::run \\Mac\Host\Volumes\Mack\汉化\——.txt
^+#!=::run \\Mac\Host\Volumes\KINGSTON\漫画
^+#!\:: ;文本文档记录测试
{
FileAppend, Another line.`n, \\Mac\Home\Documents\Test.txt
; 下面的例子使用 延续片段 来提高可读性和可维护性:
FileAppend,
(
A line of text.
By default, the hard carriage return (Enter) between the previous line and this one will be written to the file.
    This line is indented with a tab; by default, that tab will also be written to the file.
Variable references such as %Var% are expanded by default.

), \\Mac\Home\Documents\My File.txt
Return
}
;Microsoft Office Word内快捷键
#IfWinActive, ahk_class OpusApp, 填
Esc::Exit
^+#!K:: ;Word填字
{
WinGetTitle, Title ;获取窗口名
FormatTime, TimeStringStart, yyyy/MM/dd hh:mm:ss tt R
SetKeyDelay, 100
Loop, 22
{
Send ^+{Down} ;选择下一段落
Send ^c ;复制
Send {Down} ;下
Send #3 ;切换到记事本
Sleep 1000 ;延时1秒
Send ^a ;全选
Send ^v ;粘贴
Send ^s ;保存
Send #4 ;切换到PS
Sleep 1000 ;延时1秒
Send {f10} ;运行脚本
Sleep 1000 ;延时1秒
Loop ;判断脚本是否执行完
    {
        Sleep, 1000
        IfExist, \\Mac\Host\Volumes\Mack\汉化\-.txt
            break
    }
FileDelete, \\Mac\Host\Volumes\Mack\汉化\-.txt ;删除小文档
Sleep 1000 ;延时1秒
Send ^{Tab} ;切换到下一页
Sleep 1000 ;延时1秒
Send #2 ;切换到WORD
Sleep 1000 ;延时1秒
}
FileDelete, \\Mac\Home\Documents\填字完成.txt ;删除填字完成文档
Sleep 1000 ;延时1秒
SoundBeep, 750, 500 ;以较高的音高进行发音并持续半秒.
Sleep 1000 ;延时1秒
FormatTime, TimeStringEnd, yyyy/MM/dd hh:mm:ss tt R
FileAppend,
(

填字项目：%Title%
开始时间：%TimeStringStart%
完成时间：%TimeStringEnd%

), \\Mac\Home\Documents\填字完成.txt
return
}
;文件夹内快捷键
#IfWinActive, ahk_class CabinetWClass, ——
^+#!K:: ;0 day week预览截图
{
Send ^a ;全选
Sleep 1000 ;延时1秒
Send {Delete} ;删除
Sleep 2000 ;延时2秒
Send {Enter} ;是（回车）
Sleep 3000 ;延时3秒
Send {Backspace} ;上级菜单（后退）
Sleep 1000 ;延时1秒
Send {Down} ;下
Sleep 100 ;延时0.1秒
Send {Enter} ;进入（回车）
Sleep 100 ;延时0.1秒
Send {Down} ;下
Sleep 100 ;延时0.1秒
Send {Up} ;上
Sleep 100 ;延时0.1秒
Loop, 19 ;19次Shift+下
{
Send +{Down}
Sleep 100 ;延时0.1秒
}
Send ^x ;剪切
Sleep 100 ;延时0.1秒
Send {Backspace} ;上级菜单（后退）
Sleep 1000 ;延时1秒
Send {Up} ;上
Sleep 100 ;延时0.1秒
Send {Enter} ;进入（回车）
Sleep 100 ;延时0.1秒
Send ^v ;粘贴
return
}
#IfWinActive, ahk_class Chrome_WidgetWin_1 ;360浏览器内快捷键
^+#!~:: ;记录当前窗口名
WinGetTitle, Title ;获取
MsgBox 当前窗口名是%Title%
return
^+#!K:: ;从appformac网上下载App
{
CoordMode, Mouse, Client
Sleep 1000 ;延时1秒
Click ;点击
Sleep 16000 ;延时16秒
Click 2686, 208 ;跳过广告
Sleep 16000 ;延时16秒
Click 1551, 195 ;下载
Sleep 16000 ;延时16秒
Send ^w ;关闭
Sleep 1000 ;延时1秒
Send ^w ;关闭
Sleep 1000 ;延时1秒
Send ^{Tab} ;切换到下一页
return
}
#IfWinActive, ahk_class AcrobatSDIWindow ;Adobe Acrobat Pro DC内快捷键
>!h:: ;阅读模式
{
Sleep 100 ;延时0.1秒
Send ^h
return
}
>!l:: ;全屏模式
{
Sleep 100 ;延时0.1秒
Send ^l
return
}
#IfWinActive, ahk_class MangaMeeya ;MangaMeeya内快捷键
e::run \\Mac\Home\Documents\Shared Applications\欧路词典 (Mac).exe
b::run \\Mac\Home\Documents\Shared Applications\Boson (Mac).exe
f10::run \\Mac\Home\Documents\Shared Applications\Boson (Mac).exe
f11::run \\Mac\Home\Documents\Shared Applications\欧路词典 (Mac).exe
f12::run \\Mac\Home\Documents\Shared Applications\Boson (Mac).exe
^+#!~:: ;记录当前漫画进度
WinGetTitle, Title ;获取窗口名
FormatTime, Now, yyyy/MM/dd hh:mm:ss tt R
Hook = Scale
StringGetPos, position, Title, %Hook%
StringLeft, OutputVar, Title, position -2 ; 保存字符串到 OutputVar.
FileAppend,
(
项目：%OutputVar%
时间：%Now%


), \\Mac\Home\我的坚果云\漫画进度.txt
return
#IfWinActive, ahk_class FineReader12MainWindowClass
^+#!K:: ;加载
{
Send ^+o ;选项
Sleep 4000 ;延时4秒
Send !l ;从文件加载
Sleep 2000 ;延时2秒
Send y ;是
Sleep 2000 ;延时2秒
Send a ;文件名
Sleep 2000 ;延时2秒
Send {Down} ;下
Sleep 2000 ;延时2秒
Send !o ;打开
Sleep 2000 ;延时2秒
Send {Enter} ;关闭
return
}