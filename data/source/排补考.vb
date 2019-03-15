Option Compare Database
Option Explicit
Option Base 1

Private Sub Form_Load()
DoCmd.Restore
End Sub

Function Update_xsYxk(kch As String, bjh As String, iStu As Integer, iXh As Integer) As Boolean
Dim str As String
Dim rst As DAO.Recordset
Dim i As Integer
str = "select * from xsYxk where 课程号='" & kch & "' and 班级号='" & bjh & "' and 序号=1 order by 学号"
Set rst = CurrentDb.OpenRecordset(str)
Do While Not rst.EOF
    rst.Edit
    rst("序号") = iXh
    rst.Update
    i = i + 1
    If i = iStu Then Exit Do
    rst.MoveNext
Loop
rst.Close
Set rst = Nothing
Update_xsYxk = True
End Function

Private Sub 排监考_Click()
DoCmd.Close , , acSaveYes
'把jsYjkr中的监考教师清空
CurrentDb.Execute ("UPDATE jsYjkr SET jsYjkr.T1 = Null, jsYjkr.T2 = Null")

If Not testTable("teacher") Then
    MsgBox ("监考老师表不存在，请导入！")
    Exit Sub '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''退出
End If

Dim a As Integer
Dim b As Integer

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim iMaxCount As Integer    '最大的一场考试的场次数
Dim iAvg As Integer    '先非听力老师的平均监考场次数，后听力老师的平均监考场次数
Dim iCount As Integer   '总的考试时间段数
Dim iTeacher As Integer '总的监考老师数
Dim iListen As Integer  '总的听力老师数

Dim str As String
Dim rs As DAO.Recordset
Dim rst As DAO.Recordset

Dim Tdw() As String         '教师单位
Dim Tnm() As String         '教师姓名
Dim Tcount() As Byte        '教师总监考场数

str = "SELECT jsYjkr.场次标识, '' AS 备注 INTO ccYsj " _
        & "FROM jsYjkr GROUP BY jsYjkr.场次标识, '', ''"
'检查“场次与时间”表ccYsj是否存在
If Not testTable("ccYsj") Then
    CurrentDb.Execute (str)
    CurrentDb.Execute ("alter table ccYsj add column 考试时间 datetime")
ElseIf MsgBox("“场次与时间”表已经存在，是否重建并重新录入考试时间？", vbYesNo, "询问") = vbYes Then
    If CurrentData.AllTables("ccYsj").IsLoaded Then DoCmd.Close acTable, "ccYsj", acSaveNo
    CurrentDb.Execute ("drop table ccYsj")
    CurrentDb.Execute (str)
    CurrentDb.Execute ("alter table ccYsj add column 考试时间 datetime")
Else
    If CurrentData.AllTables("ccYsj").IsLoaded Then DoCmd.Close acTable, "ccYsj", acSaveNo
    DoCmd.Rename "OldccYsj", acTable, "ccYsj"
    CurrentDb.Execute (str)
    CurrentDb.Execute ("alter table ccYsj add column 考试时间 datetime")
    str = "UPDATE ccYsj INNER JOIN OldccYsj ON ccYsj.场次标识 = OldccYsj.场次标识 SET ccYsj.考试时间 = [OldccYsj].[考试时间]"
    CurrentDb.Execute (str)
    DoCmd.DeleteObject acTable, "OldccYsj"
End If

If DLookup("场次标识", "ccYsj", "考试时间 is null") Then
    MsgBox "请注意，还有考试时间没有录入！"
    Exit Sub
End If
    

iTeacher = DCount("*", "teacher", "听力 = False")
str = "SELECT teacher.学院, teacher.老师 FROM teacher WHERE ((teacher.听力 = False)) GROUP BY teacher.学院, teacher.老师 ORDER BY teacher.学院, teacher.老师"
Set rs = CurrentDb.OpenRecordset(str)
rs.MoveLast
rs.MoveFirst
j = rs.RecordCount
If iTeacher <> j Then
    MsgBox "监考老师表有重复的非听力老师，请检查！"
    rs.Close
    Set rs = Nothing
    Exit Sub '退出
End If

iCount = DMax("场次标识", "jsYjkr")

Dim kssj() As String
Dim kscc() As Byte

ReDim kssj(iCount)
ReDim kscc(iCount)

ReDim Tdw(iTeacher)
ReDim Tnm(iTeacher)
ReDim Tcount(iTeacher)

Randomize
i = 0
While Not rs.EOF    '对非听力监考老师赋初值
    i = i + 1                                           '经
    j = Int(((iTeacher - i + 1) * Rnd) + 1)             '
    k = 0                                               '
    a = 0                                               '
    Do While k < j                                      '
        a = a + 1                                       '
        If Tnm(a) = "" Then k = k + 1                   '
    Loop                                                '典
    
    Tdw(a) = rs("学院")
    Tnm(a) = rs("老师")
    Tcount(a) = 0
    CurrentDb.Execute ("update teacher set ID=" & a & " where 学院='" & Tdw(a) & "' and 老师='" & Tnm(a) & "'")
    rs.MoveNext
Wend

i = 0
a = 0
Set rs = CurrentDb.OpenRecordset("SELECT 考试时间 FROM ccYsj ORDER BY 考试时间")
Do While Not rs.EOF
    i = i + 1
    kssj(i) = Trim(rs("考试时间"))
    '取出该考试时间学分制用教室数
'    j = DCount("jsYjkr.教室号", "学分制考试安排最后结果", "ccYsj.考试时间=#" & kssj(i) & "#")
    '不使用查询“学分制考试安排最后结果”，改用SQL语句来做
    str = "SELECT DISTINCT jsYjkr.教室号 " _
& "FROM ccYsj INNER JOIN jsYjkr ON ccYsj.场次标识 = jsYjkr.场次标识 " _
& "WHERE (((ccYsj.考试时间) = #" & kssj(i) & "#))"
    Set rst = CurrentDb.OpenRecordset(str)
    rst.MoveLast
    rst.MoveFirst
    j = rst.RecordCount
    
    kscc(i) = j
    a = a + kscc(i)
    If iMaxCount < kscc(i) Then iMaxCount = kscc(i)
    rs.MoveNext
Loop
iAvg = Int(2 * a / iTeacher) ' - iCount)

Dim kch() As String
Dim jsh() As String

ReDim kch(iCount, iMaxCount)
ReDim jsh(iCount, iMaxCount)

For i = 1 To iCount
    str = "SELECT ccYsj.考试时间, jsYjkr.教室号, jsYjkr.T1, jsYjkr.T2 " _
& "FROM ccYsj INNER JOIN jsYjkr ON ccYsj.场次标识 = jsYjkr.场次标识 " _
& "GROUP BY ccYsj.考试时间, jsYjkr.教室号, jsYjkr.T1, jsYjkr.T2 " _
& "HAVING (((ccYsj.考试时间)=#" & kssj(i) & "#))"
    Set rs = CurrentDb.OpenRecordset(str)
    j = 0
    Do While Not rs.EOF
        j = j + 1
        'kch(i, j) = rs("课程号")
        jsh(i, j) = rs("教室号")
        rs.MoveNext
    Loop
    
'    Set rs = CurrentDb.OpenRecordset("select * from NotXfzResult where 考试时间=#" & kssj(i) & "#")
'    Do While Not rs.EOF
'        j = j + 1
'        jsh(i, j) = rs("教室号")
'        rs.MoveNext
'    Loop
Next i

''''''''''''''''''''''''''''''''''''''''开始排监考老师
Dim iQH As Integer, iQT As Integer '定义正在使用的教师队列
Dim strTU As String
Dim TQ() As Integer
ReDim TQ(2 * iTeacher)
For i = 1 To iTeacher
    TQ(i) = i
Next i

k = 1
For i = 1 To iCount '按时间从第一场到最后一场排
    strTU = ";"
    For j = 1 To kscc(i)
        If i = 1 And j = 1 Then iQH = iQH + 1 '第一次进队列
        
        'If Len(kch(i, j)) <> 0 Then '学分制的
        If Tcount(TQ(k)) = 0 Then iQT = iQT + 1 '进队
        'If Tcount(TQ(k)) = iAvg + 1 Then
        CurrentDb.Execute ("update jsYjkr set T1='" & TQ(k) & "' where  教室号='" & jsh(i, j) & "' and 场次标识=" & i)
        strTU = strTU & TQ(k) & ";"
        Tcount(TQ(k)) = Tcount(TQ(k)) + 1
        If Tcount(TQ(k)) = iAvg Then iQH = iQH + 1  '出队
        'Debug.Print strTU
        
        k = k + 1
        If k > iTeacher Then
            For a = iTeacher To 1 Step -1
                If InStr(strTU, ";" & a & ";") = 0 And Tcount(TQ(a)) = iAvg Then
                    iQT = iQT + 1
                    TQ(k) = a
                    Exit For
                End If
            Next a
        End If
        
'        If DLookup("语音室", "kcYrs", "课程号='" & kch(i, j) & "'") Then
'            B = B + 1 '记录听力考试的场次数,后面排听力考试的监考老师时要用到
'        Else
        If Tcount(TQ(k)) = 0 Then iQT = iQT + 1 '进队
        'If Tcount(TQ(k)) < iAvg + 1 Then
        CurrentDb.Execute ("update jsYjkr set T2='" & TQ(k) & "' where 教室号='" & jsh(i, j) & "' and 场次标识=" & i)
        strTU = strTU & TQ(k) & ";"
        Tcount(TQ(k)) = Tcount(TQ(k)) + 1
        If Tcount(TQ(k)) = iAvg Then iQH = iQH + 1 '出队
        'Debug.Print strTU
        
        k = k + 1
        If k > iTeacher Then
            For a = iTeacher To 1 Step -1
                If InStr(strTU, ";" & a & ";") = 0 And Tcount(TQ(a)) = iAvg Then
                    iQT = iQT + 1
                    TQ(k) = a
                    Exit For
                End If
            Next a
        End If
'        End If
        'End If
    Next j
    If iQH <= iTeacher Then k = iQH  '找到队头
Next i

MsgBox "监考人员顺利排完"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''以下开始汇总
End Sub

Private Sub 排教室_Click()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim iCount As Integer '场次数
Dim iClassroom As Integer '总的教室数

Dim str As String
Dim rs As DAO.Recordset


iClassroom = DCount("*", "教室", "考试否")

Dim jNo() As String '教室号
Dim jSize() As Integer '教室的考试容量
Dim jOk() As Integer '教室已排
Dim jType() As String '教室类型 1普通，2多媒体，3语音室


ReDim jNo(iClassroom)
ReDim jSize(iClassroom)
ReDim jOk(iClassroom)
ReDim jType(iClassroom)
'按使用优先级取出教室号及定员值
str = "select 教室号,教室定员,教室类型号 from 教室 where 考试否 order by 使用优先级"
Set rs = CurrentDb.OpenRecordset(str)
rs.MoveLast
rs.MoveFirst
Do While Not rs.EOF
    i = i + 1
    jNo(i) = Trim(rs("教室号"))
    jSize(i) = rs("教室定员") / 2
    jType(i) = rs("教室类型号")
    rs.MoveNext
Loop

If Not testTable("kcYrs") Then
    MsgBox "表kcYrs不存在，请检查！"
    Exit Sub
End If

If testTable("kcYbj") Then
    '判断该表是否打开，若打开则强行关闭并删除
    If CurrentData.AllTables("kcYbj").IsLoaded Then DoCmd.Close acTable, "kcYbj", acSaveNo
    CurrentDb.Execute ("drop table kcYbj")
End If

Dim tKs As Date
tKs = DLookup("ksxq", "Para") '取出考试学期
'分场次按排教室
str = "SELECT kcYrs.场次标识, 学生与选课.课程号, 学生与选课.班级号, Count(学生与选课.学号) AS 班级人数 INTO kcYbj " _
& "FROM kcYrs INNER JOIN ((学生与选课 INNER JOIN 学生 ON 学生与选课.学号 = 学生.学号) INNER JOIN 班级 ON 学生.班级号 = 班级.班级号) ON kcYrs.课程号 = 学生与选课.课程号 " _
& "WHERE (((学生与选课.开课时间) = #" & tKs & "#) And ((学生与选课.选课状态) = 0 Or (学生与选课.选课状态) Is Null) And ((kcYrs.不考试) = False)) " _
& "GROUP BY kcYrs.场次标识, 学生与选课.课程号, 学生与选课.班级号 " _
& "ORDER BY kcYrs.场次标识, 学生与选课.课程号, 学生与选课.班级号"
CurrentDb.Execute (str)

'如果jsYjkr表不存在则新建，否则清空
If Not testTable("jsYjkr") Then
    str = "create table jsYjkr(场次标识 integer,课程号 text(6),班级号 text(20), 班级人数 integer, 序号 Byte, 教室号 text(6), T1 integer, T2 integer)"
    CurrentDb.Execute (str)
Else
    CurrentDb.Execute ("delete * from jsYjkr")
End If
CurrentDb.Execute ("INSERT INTO jsYjkr( 场次标识, 课程号, 班级号 ,班级人数,序号) SELECT kcYbj.场次标识 ,kcYbj.课程号, kcYbj.班级号, kcYbj.班级人数 ,1 FROM kcYbj")
'CurrentDb.Execute ("update jsYjkr set 序号=1")

'对学生与选课表里的序号进行更新，全更新为1
CurrentDb.Execute ("update xsYxk set 序号=1 where 序号<>1")

iCount = DMax("场次标识", "jsYjkr")

For i = 1 To iCount
    str = "select * from kcYbj where 场次标识=" & i & " order by 课程号" '不考虑使用语音室与否
    Set rs = CurrentDb.OpenRecordset(str)
    rs.MoveLast
    rs.MoveFirst
    'Dim iA As Integer
    'Dim iB As Integer
    'Dim m As Integer
    Dim ms As Integer '动态记录某教室中补考课程门数
    Dim ss As Integer '动态记录某教室中补考人数
    Dim kh As String
    ms = 0
    ss = 0
    kh = ""
    For j = 1 To iClassroom
        jOk(j) = 0
    Next j
    Do While Not rs.EOF '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''do1
        For j = 1 To iClassroom ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''for1
            If jOk(j) = 0 Then ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''If1
                Do While Not rs.EOF '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''do2
                    If rs("班级人数") <= jSize(j) Then '''''''''''''''''''''''''''''If2
                        If rs("课程号") <> kh Then
                            ms = ms + 1
                        End If
                        ss = ss + rs("班级人数")
                        If ss <= jSize(j) And ms <= 10 Then ''''''''''''''''''''''If3
                            CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(j) & "' where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "'")
                            kh = rs("课程号")
                            '预读
                            rs.MoveNext
                            If Not rs.EOF Then
                                If DSum("班级人数", "kcYbj", "课程号='" & Trim(rs("课程号")) & "'") + ss > jSize(j) And rs("课程号") <> kh Then '下一班级读入将超员或已读人员已满员
                                    ms = 0
                                    ss = 0
                                    kh = ""
                                    jOk(j) = 1
                                    '将虚位教室清空复位
                                    For k = 1 To iClassroom
                                        If jOk(k) = 8 Then
                                            jOk(k) = 0
                                        End If
                                    Next k
                                    Exit For '这个教室排完
                                ElseIf ss + rs("班级人数") > jSize(j) Then '下一课程全部人数读入将超员
                                    ms = 0
                                    ss = 0
                                    kh = ""
                                    jOk(j) = 1
                                    '将虚位教室清空复位
                                    For k = 1 To iClassroom
                                        If jOk(k) = 8 Then
                                            jOk(k) = 0
                                        End If
                                    Next k
                                    Exit For
                                End If ' 不回读了
                           End If
                        Else '已达教室定额或课程门数已达上限''''''''''''''ElseIf3
                            ms = 0
                            ss = 0
                            kh = ""
                            jOk(j) = 1
                             '将虚位教室清空复位
                            For k = 1 To iClassroom
                                If jOk(k) = 8 Then
                                   jOk(k) = 0
                                End If
                            Next k
                           Exit For '这个教室排完
                        End If '''''''''''''''''''''''''''''''''''''''''''EndIf3
                    Else '此教室虚位 ''''''''''''''''''esleif2
                        jOk(j) = 8
                    End If '''''''''''''''''''''''''''''end if2
                Loop ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''do2
                Exit For
            End If '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''end if1
        Next j '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''for1
        
        If j = iClassroom + 1 And Not rs.EOF Then  '这个班级人数太多,只好分多个教室来排,但注意要放在同一楼，最好是同一层。
           rs.MoveNext
           MsgBox "第" & i & "场的有课程没法排"
        End If
        
    Loop ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''do1
Next i
MsgBox "已排完教室"
End Sub

Private Sub 退出_Click()
Dim str As String
Dim rs As DAO.Recordset
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''以下开始汇总
If testTable("Total") Then
    If CurrentData.AllTables("Total").IsLoaded Then DoCmd.Close acTable, "Total", acSaveNo
    CurrentDb.Execute ("drop table Total")
End If
CurrentDb.Execute ("SELECT jsYjkr.场次标识, jsYjkr.课程号, jsYjkr.班级号, jsYjkr.班级人数, jsYjkr.序号, jsYjkr.教室号, jsYjkr.T1 AS T INTO Total FROM jsYjkr")

str = "INSERT INTO Total ( 场次标识, 课程号, 班级号, 班级人数, 序号, 教室号, T ) " _
& "SELECT jsYjkr.场次标识, jsYjkr.课程号, jsYjkr.班级号, jsYjkr.班级人数, jsYjkr.序号, jsYjkr.教室号, jsYjkr.T2 AS T " _
& "FROM jsYjkr"
CurrentDb.Execute (str)

DoCmd.Close , , acSaveYes
End Sub
