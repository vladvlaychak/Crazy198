Dim Num1, Sign, Num2
Num1 = InputBox ("First nomber:", "Calculator V1", "")
Sign = InputBox ("Symbol" & Chr (10) & "(Support: +, |, &, =, ):"_
& Chr (10) & "For take helping please write About", "Calculator V1", "")
Num2 = InputBox ("Second nomber:" & Chr (10) & "(UNessential)", "Calculator V1", "")
'[|]КОньюнкция
If Sign = "|" Then
Dim MinResult
MinResult = cint (Num1) or cint (Num2)
MsgBox MinResult, 0, "Calculator V1"
End If
'[|]КОньюнкция
If Sign = "or" Then
Dim MinResulte
MinResulte = cint (Num1) or cint (Num2)
MsgBox MinResulte, 0, "Calculator V1"
End If
'[&]Дизьюнкция
If Sign = "&" Then
Dim MulResult
MulResult = cint (Num1) and cint (Num2) 
MsgBox MulResult, 0, "Calculator V1"
End If
'[&]Дизьюнкция
If Sign = "and" Then
Dim MulResulte
MulResulte = cint (Num1) and cint (Num2) 
MsgBox MulResulte, 0, "Calculator V1"
End If
'[=]Эквивалентность
If Sign = "=" Then
Dim DivResult
DivResult = cint (Num1) Eqv cint (Num2) 
MsgBox DivResult, 0, "Calculator V1"
End If
'[+]СЛожение по модулю
If Sign = "+" Then
Dim Root
Root = cint (Num1) Xor (Num2)
MsgBox Root, 0, "Calculator V1"
End If
'[About]Информация о скрипте
If Sign = "About" Then
MsgBox "This calculator VBScript." & Chr (10) &_
"His can completed prosess such as:Conqnction, Dizunction, Implication, SummforModul."_
& Chr (10) & "(C)VladVladichak 18/09/2019", 0, "Calculator V1"
End If
'[Error]В случаи ошибок
If Num1 = "" Then
MsgBox "You are not press nomber", 0, "Calculator V1"
End If
If Sign = "" Then
MsgBox "You are not press simbol!", 0, "Calculator V1"
End If