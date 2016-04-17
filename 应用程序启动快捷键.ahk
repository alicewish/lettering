^+#!B::run C:\Users\Alicewish\AppData\Roaming\baidu\BaiduYunGuanjia\baiduyunguanjia.exe
^+#!C::run C:\Program Files (x86)\Calibre2\calibre.exe
^+#!D::run C:\Program Files (x86)\DuplicateCleanerPro\DuplicateCleaner.exe
^+#!E::run C:\Program Files\Everything\Everything.exe
^+#!F::run C:\Program Files (x86)\Mozilla Firefox\firefox.exe
^+#!H::run \\Mac\Host\Volumes\Mack\汉化
^+#!I::run C:\Program Files (x86)\Internet Download Manager\IDMan.exe
^+#!J::run C:\Program Files (x86)\按键精灵2014\按键精灵2014.exe
^+#!L::run C:\Program Files (x86)\iTools\iTools.exe
^+#!N::run C:\Program Files (x86)\ABBYY FineReader 12\FineReader.exe
^+#!O::run C:\Program Files\totalcmd\TOTALCMD64.EXE
^+#!P::run C:\Program Files (x86)\Adobe\Photoshop CS6\Photoshop.exe
^+#!Q::run C:\Program Files (x86)\Tencent\QQ\Bin\QQScLauncher.exe
^+#!R::run C:\Program Files (x86)\Advanced Renamer\ARen.exe
^+#!T::run C:\Program Files (x86)\Tongbu\Tongbu.exe
^+#!U::run C:\Program Files\UltraEdit\Uedit32.exe
^+#!X::run C:\Program Files\ComicStudio EX\Tool\CS_EX.exe
^+#!Y::run C:\Users\Alicewish\AppData\Local\Youdao\Dict\Application\YodaoDict.exe
^+#!1::run \\Mac\Dropbox
^+#!2::run C:\Users\Alicewish\Google 云端硬盘
^+#!3::run \\Mac\Home\Downloads
^+#!4::run G:\Comics
^+#!5::run G:\熟肉
^+#!6::run \\Mac\Dropbox\PSD工作流
^+#!7::run \\Mac\Home\Downloads\Compressed
^+#!8::run \\Mac\Home\Documents\Tencent Files\782841300\FileRecv
^+#!9::run \\Mac\Home\Desktop
^+#!-::run \\Mac\Host\Volumes\Mack\汉化\——.txt
#IfWinActive, ahk_class OpusApp, 填
^+#!K::
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
return
#IfWinActive, ahk_class CabinetWClass, ——
^+#!K::
Send ^a ;全选
Sleep 1000 ;延时1秒
Send {Delete} ;删除
Sleep 2000 ;延时2秒
Send {Enter} ;是（回车）
Sleep 1000 ;延时1秒
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
#IfWinActive, ahk_class Chrome_WidgetWin_1
^+#!K::
CoordMode, Mouse, Client
Sleep 1000 ;延时1秒
Click ;点击
Sleep 16000 ;延时16秒
Click 2642, 172 ;跳过广告
Sleep 16000 ;延时16秒
Click 1548, 161 ;下载
Sleep 16000 ;延时16秒
Send ^w ;关闭
Sleep 1000 ;延时1秒
Send ^w ;关闭
Sleep 1000 ;延时1秒
Send ^{Tab} ;切换到下一页
return
