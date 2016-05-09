Sub 中英文标点互换()
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

Sub 着重修正()
Dim BeforeChange() As Variant, AfterChange() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim N As Byte
BeforeChange = Array("|——", "|—", "|,", "|!", "|?!", "|?", "|…", "|。”", "|。", "|”", "“|")
AfterChange = Array("——|", "—|", ",|", "!|", "?!|", "?|", "…|", "。”|", "。|", "”|", "|“")
myArray1 = BeforeChange
myArray2 = AfterChange
Application.ScreenUpdating = False '关闭屏幕更新
For N = 0 To UBound(BeforeChange) '从数组的下标到上标间作一个循环
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = False '不使用通配符
'查找替换
.Execute findtext:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
.MatchWildcards = False
End With
Next
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

Sub 填字符号转换()
Dim BeforeChange() As Variant, AfterChange() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim N As Byte
BeforeChange = Array("--|",",|",":|","?!|","!|","?|","|——","|……","…|",".""|",".|","·","|--","|…","|""","——|","—|","，|","：|","？！|","！|","？|","……|","。”|","。|","”|","|—","|“","、|")
AfterChange = Array(鲆|","鲡|","鲛|","鲧|","鲚|","鲩|","|鲮","|鳆","鲠|","鲋|","鲴|","鲂","|鲆","|鲠","|鲕","鲮|","鲵|","鲲|","鲒|","鲎|","鲣|","鳇|","鳆|","鲱|","鲼|","鲅|","|鲵","|鲽","鳠|)
myArray1 = BeforeChange
myArray2 = AfterChange
Application.ScreenUpdating = False '关闭屏幕更新
For N = 0 To UBound(BeforeChange) '从数组的下标到上标间作一个循环
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = False '不使用通配符
'查找替换
.Execute findtext:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
.MatchWildcards = False
End With
Next
Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub 填字符号转回()
Dim BeforeChange() As Variant, AfterChange() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim N As Byte
BeforeChange = Array("--|",",|",":|","?!|","!|","?|","|——","|……","…|",".""|",".|","·","|--","|…","|""","——|","—|","，|","：|","？！|","！|","？|","……|","。”|","。|","”|","|—","|“","、|")
AfterChange = Array(鲆|","鲡|","鲛|","鲧|","鲚|","鲩|","|鲮","|鳆","鲠|","鲋|","鲴|","鲂","|鲆","|鲠","|鲕","鲮|","鲵|","鲲|","鲒|","鲎|","鲣|","鳇|","鳆|","鲱|","鲼|","鲅|","|鲵","|鲽","鳠|)
myArray1 = AfterChange
myArray2 = BeforeChange
Application.ScreenUpdating = False '关闭屏幕更新
For N = 0 To UBound(BeforeChange) '从数组的下标到上标间作一个循环
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = False '不使用通配符
'查找替换
.Execute findtext:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
.MatchWildcards = False
End With
Next
Application.ScreenUpdating = True '恢复屏幕更新
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

Sub 对k的格式修正()
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

Sub ChangeCAPStoBold()
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


