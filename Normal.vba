Sub ToggleInterpunction() '中英文标点互换
Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim msgResult As VbMsgBoxResult, N As Byte
'定义一个中文标点的数组对象
ChineseInterpunction = Array("、", "。", "，", "；", "：", "？", "！", "……", "—", "～", "（", "）", "《", "》")
'定义一个英文标点的数组对象
EnglishInterpunction = Array("、", ".", ",", ";", ":", "?", "!", "…", "-", "~", "(", ")", "&lt;", "&gt;")
'提示用户交互的MSGBOX对话框
msgResult = MsgBox("您想中英标点互换吗?按Y将中文标点转为英文标点,按N将英文标点转为中文标点!", vbYesNoCancel)
Select Case msgResult
Case vbCancel
Exit Sub '如果用户选择了取消按钮,则退出程序运行
Case vbYes '如果用户选择了YES,则将中文标点转换为英文标点
myArray1 = ChineseInterpunction
myArray2 = EnglishInterpunction
strFind = "“(*)”"
strRep = """\1"""
Case vbNo '如果用户选择了NO,则将英文标点转换为中文标点
myArray1 = EnglishInterpunction
myArray2 = ChineseInterpunction
strFind = """(*)"""
strRep = "“\1”"
End Select
Application.ScreenUpdating = False '关闭屏幕更新
For N = 0 To UBound(ChineseInterpunction) '从数组的下标到上标间作一个循环
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = False '不使用通配符
'查找相应的英文标点,替换为对应的中文标点
.Execute findtext:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
.MatchWildcards = False
End With
Next
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = True '使用通配符
.Execute findtext:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
.MatchWildcards = False
End With
Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub ChangeInterpunction() '中英文标点互换改
Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim msgResult As VbMsgBoxResult, N As Byte
'定义一个中文标点的数组对象
ChineseInterpunction = Array("、", "，", "；", "？", "！", "……", "～", "（", "）", "《", "》")
'定义一个英文标点的数组对象
EnglishInterpunction = Array("、", ",", ";", "?", "!", "…", "~", "(", ")", "&lt;", "&gt;")
'提示用户交互的MSGBOX对话框
msgResult = MsgBox("您想中英标点互换吗?按Y将中文标点转为英文标点,按N将英文标点转为中文标点!", vbYesNoCancel)
Select Case msgResult
Case vbCancel
Exit Sub '如果用户选择了取消按钮,则退出程序运行
Case vbYes '如果用户选择了YES,则将中文标点转换为英文标点
myArray1 = ChineseInterpunction
myArray2 = EnglishInterpunction
Case vbNo '如果用户选择了NO,则将英文标点转换为中文标点
myArray1 = EnglishInterpunction
myArray2 = ChineseInterpunction
End Select
Application.ScreenUpdating = False '关闭屏幕更新
For N = 0 To UBound(ChineseInterpunction) '从数组的下标到上标间作一个循环
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = False '不使用通配符
'查找相应的英文标点,替换为对应的中文标点
.Execute findtext:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
.MatchWildcards = False
End With
Next
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = True '使用通配符
.Execute findtext:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
.MatchWildcards = False
End With
Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub 测试()
Selection.Find.Replacement.Font.Color = -738131969
With Selection.Find
.Text = "([0-9])([0-9])"
.Replacement.Text = "^p\1\2"
.Wrap = wdFindContinue
.Format = True
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 校对转填字()
With Selection.Find
.Text = "^l"
.Replacement.Text = "^p"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "([0-9])([0-9])^13^13"
.Replacement.Text = ""
.Wrap = wdFindContinue
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "^p^p"
.Replacement.Text = "^p"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "^p"
.Replacement.Text = "^p^p"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "^p^p^p^p^p"
.Replacement.Text = "^p"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 填字转分页()
With Selection.Find
.Text = "^p"
.Replacement.Text = "^l"
.Wrap = wdFindContinue
.MatchByte = True
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "^11^11([0-9])([0-9])^11^11"
.Replacement.Text = "^p^l"
.Wrap = wdFindContinue
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 着重修正()
With Selection.Find
.Text = "|——"
.Replacement.Text = "——|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|—"
.Replacement.Text = "—|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|,"
.Replacement.Text = ",|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|!"
.Replacement.Text = "!|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|?!"
.Replacement.Text = "?!|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|?"
.Replacement.Text = "?|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|…"
.Replacement.Text = "…|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|。”"
.Replacement.Text = "。”|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|。"
.Replacement.Text = "。|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|”"
.Replacement.Text = "”|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "—|"
.Replacement.Text = "|—"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "——|"
.Replacement.Text = "|——"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "…|"
.Replacement.Text = "|…"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "“|"
.Replacement.Text = "|“"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 填字符号转换()
With Selection.Find
.Text = "--|"
.Replacement.Text = "鲆|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = ",|"
.Replacement.Text = "鲡|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = ":|"
.Replacement.Text = "鲛|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "?!|"
.Replacement.Text = "鲧|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "!|"
.Replacement.Text = "鲚|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "?|"
.Replacement.Text = "鲩|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|——"
.Replacement.Text = "|鲮"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|……"
.Replacement.Text = "|鳆"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "…|"
.Replacement.Text = "鲠|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = ".""|"
.Replacement.Text = "鲋|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = ".|"
.Replacement.Text = "鲴|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "·"
.Replacement.Text = "鲂"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|--"
.Replacement.Text = "|鲆"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|…"
.Replacement.Text = "|鲠"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|"""
.Replacement.Text = "|鲕"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "——|"
.Replacement.Text = "鲮|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "—|"
.Replacement.Text = "鲵|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "，|"
.Replacement.Text = "鲲|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "：|"
.Replacement.Text = "鲒|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "？！|"
.Replacement.Text = "鲎|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "！|"
.Replacement.Text = "鲣|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "？|"
.Replacement.Text = "鳇|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "……|"
.Replacement.Text = "鳆|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "。”|"
.Replacement.Text = "鲱|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "。|"
.Replacement.Text = "鲼|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "”|"
.Replacement.Text = "鲅|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|—"
.Replacement.Text = "|鲵"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|“"
.Replacement.Text = "|鲽"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "、|"
.Replacement.Text = "鳠|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 填字符号校正()
With Selection.Find
.Text = "鲆|"
.Replacement.Text = "--|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲡|"
.Replacement.Text = ",|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲛|"
.Replacement.Text = ":|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲧|"
.Replacement.Text = "?!|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲚|"
.Replacement.Text = "!|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲩|"
.Replacement.Text = "?|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|鲮"
.Replacement.Text = "|——"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|鳆"
.Replacement.Text = "|……"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲠|"
.Replacement.Text = "…|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲋|"
.Replacement.Text = ".""|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲴|"
.Replacement.Text = ".|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲂"
.Replacement.Text = "·"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|鲆"
.Replacement.Text = "|--"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|鲠"
.Replacement.Text = "|…"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|鲕"
.Replacement.Text = "|"""
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲮|"
.Replacement.Text = "——|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲵|"
.Replacement.Text = "—|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲲|"
.Replacement.Text = "，|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲒|"
.Replacement.Text = "：|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲎|"
.Replacement.Text = "？！|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲣|"
.Replacement.Text = "！|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鳇|"
.Replacement.Text = "？|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鳆|"
.Replacement.Text = "……|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲱|"
.Replacement.Text = "。”|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲼|"
.Replacement.Text = "。|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鲅|"
.Replacement.Text = "”|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|鲵"
.Replacement.Text = "|—"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "|鲽"
.Replacement.Text = "|“"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "鳠|"
.Replacement.Text = "、|"
.Wrap = wdFindContinue
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 缩行()
With Selection.Find
.Text = "^p"
.Replacement.Text = "の"
.Wrap = wdFindContinue
.MatchByte = True
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 世图()
With Selection.Find
.Text = "^p"
.Replacement.Text = "^p^p"
.Wrap = wdFindContinue
.MatchByte = True
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "^p^p^p^p"
.Replacement.Text = "^p^p^p"
.Wrap = wdFindContinue
.MatchByte = True
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "([0-9])([0-9])"
.Replacement.Text = "0\1\2"
.Wrap = wdFindContinue
End With
Selection.Find.Execute Replace:=wdReplaceAll
Selection.Find.Replacement.Font.Color = -738131969
With Selection.Find
.Text = "([0-9])([0-9])([0-9])"
.Replacement.Text = ""
.Wrap = wdFindContinue
.Format = True
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "([0-9])([0-9])([0-9])"
.Replacement.Text = "^p\1\2\3"
.Wrap = wdFindContinue
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub k() '对k的格式修正
With Selection.Find
.Text = " "
.Replacement.Text = "^p"
.Wrap = wdFindContinue
.MatchByte = True
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = vbTab
.Replacement.Text = "^p"
.Wrap = wdFindContinue
.MatchByte = True
.MatchWildcards = False
End With
Selection.Find.Execute Replace:=wdReplaceAll
With Selection.Find
.Text = "【([一-龥]{3,5})】"
.Replacement.Text = "^13"
.Wrap = wdFindContinue
.Format = True
End With
Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub ChangeCAPStoBold() '将全大写单词转为粗体
With Selection.Find
    .Text = "(<[A-Z.]{2,})"
    .Replacement.Text = "\1"
    .Replacement.Font.Bold = True
        .Wrap = wdFindContinue
    .Format = True
End With
While Selection.Find.Execute
   Selection.Range.Case = wdTitleWord
   Selection.Font.Bold = True
Wend
End Sub

