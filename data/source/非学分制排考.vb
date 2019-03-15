Option Compare Database
Option Explicit
Option Base 1

Private Sub 确定_Click()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim ij As Integer
Dim iA As Integer
Dim iB As Integer
Dim ccs As Integer
Dim str As String
Dim rs As DAO.Recordset
Dim rsA As DAO.Recordset
Dim cc() As String
Dim sj() As String
Dim ok() As Boolean         '非学分制下某一场次的某个教室是否已排的标志

Dim jNo() As String         '教室号
Dim jSize() As Integer      '教室的考试容量
Dim jOk() As Byte           '教室已排
Dim jType() As String    '教室类型 1普通，2多媒体，3语音室

Dim bok As Boolean

Set rs = CurrentDb.OpenRecordset("SELECT distinct 场次标识 FROM AllccYjs WHERE 已用否=False")
rs.MoveLast
rs.MoveFirst
ccs = rs.RecordCount
ReDim cc(ccs)
ReDim sj(ccs)
ReDim ok(ccs)   '这一场次是否已排

If Not testTable("NotXfz") Then
    MsgBox ("非学分制的考试安排表不存在，请导入！")
    Exit Sub
Else
    str = "select distinct 考试时间 from NotXfz"
    Set rs = CurrentDb.OpenRecordset(str)
    rs.MoveLast
    rs.MoveFirst
    i = rs.RecordCount
    If i > ccs Then
        MsgBox "考试的时间段太多，请检查并压缩！"
        Exit Sub
    End If
End If
i = 0

'生成“非学分制考试安排结果表”并导入初始数据
If Not testTable("NotXfzResult") Then
    CurrentDb.Execute ("select * into NotXfzResult from NotXfz")
    CurrentDb.Execute ("delete from NotXfzResult")
    CurrentDb.Execute ("alter table NotXfzResult add column 序号 Byte, 教室号 text(8),T1 integer,T2 integer")
Else
    If CurrentData.AllTables("NotXfzResult").IsLoaded Then DoCmd.Close acTable, "NotXfzResult", acSaveNo
    CurrentDb.Execute ("delete from NotXfzResult")
End If

'对剩余教室按数量与总容量排序（降序）
str = "SELECT AllccYjs.场次标识, Count(AllccYjs.教室号) AS 教室数, Sum(教室.教室定员) AS 总容量 " _
        & "FROM AllccYjs INNER JOIN 教室 ON AllccYjs.教室号 = 教室.教室号 " _
        & "WHERE (((AllccYjs.已用否) = False)) " _
        & "GROUP BY AllccYjs.场次标识 " _
        & "ORDER BY Count(AllccYjs.教室号) DESC , Sum(教室.教室定员) DESC"
Set rs = CurrentDb.OpenRecordset(str)
While Not rs.EOF
    i = i + 1
    cc(i) = rs("场次标识")
    sj(i) = DLookup("考试时间", "ccYsj", "场次标识=" & rs("场次标识"))  '?
    ok(i) = False
    rs.MoveNext
Wend
'对考试的时间段按考试课程与考试总人数排序（降序）
str = "SELECT NotXfz.考试时间, Count(NotXfz.考试课程) AS 课程门数, Sum(NotXfz.人数) AS 考试人数 " _
        & "FROM NotXfz " _
        & "GROUP BY NotXfz.考试时间 " _
        & "ORDER BY Count(NotXfz.考试课程) DESC , Sum(NotXfz.人数) DESC"
Set rs = CurrentDb.OpenRecordset(str)
rs.MoveLast
rs.MoveFirst
i = rs.RecordCount
Dim NotXfzSj() As Boolean
ReDim NotXfzSj(i)

Dim m As Integer
Dim n As Boolean
While Not rs.EOF
    n = True
    m = m + 1
    i = DLookup("场次标识", "ccYsj", "考试时间=#" & rs("考试时间") & "#")   '找出对应是哪一场
    j = DCount("*", "AllccYjs", "已用否=false and 场次标识=" & i)   '找出这一场剩余的教室数
    ReDim jNo(j)
    ReDim jSize(j)
    ReDim jOk(j)
    ReDim jType(j)
    For k = 1 To j  '清零
        jNo(k) = ""
        jOk(k) = 0
        jSize(k) = 0
        jType(k) = ""
    Next k

    str = "SELECT 教室.教室号, 教室.教室类型号, 教室.教室定员 " _
            & "FROM AllccYjs INNER JOIN 教室 ON AllccYjs.教室号 = 教室.教室号 " _
            & "WHERE (((AllccYjs.已用否)=False) AND ((AllccYjs.场次标识)=" & i & ")) " _
            & "ORDER BY 教室.使用优先级" '教室.教室定员,
    Set rsA = CurrentDb.OpenRecordset(str)
    rsA.MoveLast
    rsA.MoveFirst
    k = 0
    Do While Not rsA.EOF   '取出各教室的信息
        k = k + 1
        jNo(k) = Trim(rsA("教室号"))
        jSize(k) = rsA("教室定员") / 2
        jType(k) = rsA("教室类型号")
        rsA.MoveNext
    Loop

    str = "select * from NotXfz where 考试时间=#" & Trim(rs("考试时间")) & "# order by 人数 desc"
    Set rsA = CurrentDb.OpenRecordset(str)
    Do While Not rsA.EOF
        bok = False
        If rsA("人数") > 120 Then
            iA = rsA("人数") / 3
            iB = rsA("人数") - iA * 2
            For ij = 1 To j - 2
                If iA <= jSize(ij) And jOk(ij) = 0 And Not bok Then
                    For k = ij + 1 To j - 1
                        If iA <= jSize(k) And jOk(k) = 0 And Not bok Then
                            For l = k + 1 To j
                                If iB <= jSize(l) And jOk(l) = 0 And Left(jNo(l), 1) = Left(jNo(k), 1) Then
                                    CurrentDb.Execute ("insert into NotXfzResult (单位,考试课程,人数,考试时间,班级,教室号,序号) values ('" _
                                    & rsA("单位") & "','" & rsA("考试课程") & "'," & iA & ",'" & rsA("考试时间") & "','" & rsA("班级") & "','" & jNo(ij) & "',1)")
                                    jOk(ij) = 1
                                    CurrentDb.Execute ("insert into NotXfzResult (单位,考试课程,人数,考试时间,班级,教室号,序号) values ('" _
                                    & rsA("单位") & "','" & rsA("考试课程") & "'," & iA & ",'" & rsA("考试时间") & "','" & rsA("班级") & "','" & jNo(k) & "',2)")
                                    jOk(k) = 1
                                    CurrentDb.Execute ("insert into NotXfzResult (单位,考试课程,人数,考试时间,班级,教室号,序号) values ('" _
                                    & rsA("单位") & "','" & rsA("考试课程") & "'," & iB & ",'" & rsA("考试时间") & "','" & rsA("班级") & "','" & jNo(l) & "',3)")
                                    jOk(l) = 1
                                    bok = True
                                    Exit For '这个班级排定
                                End If
                            Next l
                        End If
        
                        If bok Then Exit For
                    Next k
                End If
                
                If bok Then Exit For
            Next ij
        ElseIf rsA("人数") > 60 Then
            iA = rsA("人数") / 2
            iB = rsA("人数") - iA
            For k = 1 To j - 1
                If iA <= jSize(k) And jOk(k) = 0 And Not bok Then
                    For l = k + 1 To j
                        If iB <= jSize(l) And jOk(l) = 0 And Left(jNo(l), 1) = Left(jNo(k), 1) Then
                            CurrentDb.Execute ("insert into NotXfzResult (单位,考试课程,人数,考试时间,班级,教室号,序号) values ('" _
                            & rsA("单位") & "','" & rsA("考试课程") & "'," & iA & ",'" & rsA("考试时间") & "','" & rsA("班级") & "','" & jNo(k) & "',1)")
                            jOk(k) = 1
                            CurrentDb.Execute ("insert into NotXfzResult (单位,考试课程,人数,考试时间,班级,教室号,序号) values ('" _
                            & rsA("单位") & "','" & rsA("考试课程") & "'," & iB & ",'" & rsA("考试时间") & "','" & rsA("班级") & "','" & jNo(l) & "',2)")
                            jOk(l) = 1
                            bok = True
                            Exit For '这个班级排定
                        End If
                    Next l
                End If

                If bok Then Exit For
            Next k
        Else
            For k = 1 To j
                If rsA("人数") <= jSize(k) And jOk(k) = 0 Then
                    CurrentDb.Execute ("insert into NotXfzResult (单位,考试课程,人数,考试时间,班级,教室号,序号) values ('" _
                    & rsA("单位") & "','" & rsA("考试课程") & "'," & rsA("人数") & ",'" & rsA("考试时间") & "','" & rsA("班级") & "','" & jNo(k) & "',1)")
                    jOk(k) = 1
                    Exit For '这个班级排定
                End If
            Next k
        End If

        If k = j + 1 Then
            MsgBox "请注意：" & Trim(rs("考试时间")) & "的考试排不下！"
            n = False
            CurrentDb.Execute ("delete from NotXfzResult where 考试时间=#" & rsA("考试时间") & "#") '回滚
            Exit Do
        End If

        rsA.MoveNext
    Loop
    
    NotXfzSj(m) = n
    rs.MoveNext
Wend

rs.MoveFirst
i = 0
m = 0
While Not rs.EOF
    m = m + 1
    If Not NotXfzSj(m) Then
        Do
            i = i + 1   '通过cc(i)来与NotXfz中对应顺序的考试时间段挂钩
        Loop Until Not ok(i)
        j = DCount("*", "AllccYjs", "已用否=false and 场次标识=" & cc(i))   '找出这一场剩余的教室数
        ReDim jNo(j)
        ReDim jSize(j)
        ReDim jOk(j)
        ReDim jType(j)
        For k = 1 To j  '清零
            jNo(k) = ""
            jOk(k) = 0
            jSize(k) = 0
            jType(k) = ""
        Next k
    
        str = "SELECT 教室.教室号, 教室.教室类型号, 教室.教室定员 " _
                & "FROM AllccYjs INNER JOIN 教室 ON AllccYjs.教室号 = 教室.教室号 " _
                & "WHERE (((AllccYjs.已用否)=False) AND ((AllccYjs.场次标识)=" & cc(i) & ")) " _
                & "ORDER BY 教室.教室定员, 教室.教室号"
        Set rsA = CurrentDb.OpenRecordset(str)
        rsA.MoveLast
        rsA.MoveFirst
        k = 0
        Do While Not rsA.EOF   '取出各教室的信息
            k = k + 1
            jNo(k) = Trim(rsA("教室号"))
            jSize(k) = rsA("教室定员") / 2
            jType(k) = rsA("教室类型号")
            rsA.MoveNext
        Loop
    
        str = "select * from NotXfz where 考试时间=#" & Trim(rs("考试时间")) & "# order by 人数 desc"
        Set rsA = CurrentDb.OpenRecordset(str)
        While Not rsA.EOF
            bok = False
            If rsA("人数") > 130 Then
                '先不考虑人数超过120的情况
            ElseIf rsA("人数") > 64 Then
                iA = rsA("人数") / 2
                iB = rsA("人数") - iA
                For k = 1 To j - 1
                    If iA <= jSize(k) And jOk(k) = 0 And Not bok Then
                        For l = k + 1 To j
                            If iB <= jSize(l) And jOk(l) = 0 And Left(jNo(l), 1) = Left(jNo(k), 1) Then
                                CurrentDb.Execute ("insert into NotXfzResult (单位,考试课程,人数,考试时间,班级,教室号,序号) values ('" _
                                & rsA("单位") & "','" & rsA("考试课程") & "'," & iA & ",'" & sj(cc(i)) & "','" & rsA("班级") & "','" & jNo(k) & "',1)")
                                jOk(k) = 1
                                CurrentDb.Execute ("insert into NotXfzResult (单位,考试课程,人数,考试时间,班级,教室号,序号) values ('" _
                                & rsA("单位") & "','" & rsA("考试课程") & "'," & iB & ",'" & sj(cc(i)) & "','" & rsA("班级") & "','" & jNo(l) & "',2)")
                                jOk(l) = 1
                                bok = True
                                Exit For '这个班级排定
                            End If
                        Next l
                    End If
    
                    If bok Then Exit For
                Next k
            Else
                For k = 1 To j
                    If rsA("人数") <= jSize(k) And jOk(k) = 0 Then
                        CurrentDb.Execute ("insert into NotXfzResult (单位,考试课程,人数,考试时间,班级,教室号,序号) values ('" _
                        & rsA("单位") & "','" & rsA("考试课程") & "'," & rsA("人数") & ",'" & sj(i) & "','" & rsA("班级") & "','" & jNo(k) & "',1)")
                        jOk(k) = 1
                        Exit For '这个班级排定
                    End If
                Next k
            End If
    
            If k = j + 1 Then MsgBox "Terrible"
    
            rsA.MoveNext
        Wend
    End If
    rs.MoveNext
Wend
MsgBox "教室已顺利排完！"
End Sub

Private Sub 退出_Click()
On Error GoTo 退出_Click_error
DoCmd.Close , , acSaveYes
If MsgBox("是否开始排监考员操作？", vbYesNo) <> vbYes Then
    Exit Sub
End If

'把jsYjkr中的监考教师清空
CurrentDb.Execute ("UPDATE jsYjkr SET jsYjkr.T1 = Null, jsYjkr.T2 = Null")
CurrentDb.Execute ("UPDATE NotXfzResult SET NotXfzResult.T1 = Null, NotXfzResult.T2 = Null")
If Not testTable("teacher") Then
    MsgBox ("监考老师表不存在，请导入！")
    Exit Sub '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''退出
End If

Dim a As Integer
Dim b As Integer
Dim c As Integer

Dim d As Single
Dim bBack As Boolean

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim iTotal '总考场数
Dim iMaxKcs As Integer    '最大的一场考试的考场数
Dim iAvg As Integer    '先非听力老师的平均监考场次数，后听力老师的平均监考场次数
Dim iCount As Integer   '总的考试时间段数
Dim iTeacher As Integer '总的监考老师数
Dim iListen As Integer  '总的听力老师数

Dim str As String
Dim rs As DAO.Recordset

Dim Tdw() As String         '教师单位
Dim Tnm() As String         '教师姓名
Dim Tcount() As Byte        '教师总监考场数


iTeacher = DCount("*", "teacher", "听力否 = False")
str = "SELECT teacher.单位, teacher.姓名 FROM teacher WHERE ((teacher.听力否 = False)) GROUP BY teacher.单位, teacher.姓名 ORDER BY teacher.单位, teacher.姓名"
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

iCount = DMax("场次标识", "ccYsj") '20101209

Dim strKssj() As String
Dim bKcs() As Byte
ReDim strKssj(iCount)
ReDim bKcs(iCount)

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
    
    Tdw(a) = rs("单位")
    Tnm(a) = rs("姓名")
    Tcount(a) = 0
    CurrentDb.Execute ("update teacher set ID=" & a & " where 单位='" & Tdw(a) & "' and 姓名='" & Tnm(a) & "'")
    rs.MoveNext
Wend

i = 0
iTotal = 0
Set rs = CurrentDb.OpenRecordset("SELECT 考试时间 FROM ccYsj ORDER BY 考试时间")
Do While Not rs.EOF
    i = i + 1
    strKssj(i) = Trim(rs("考试时间"))
    j = DCount("jsYjkr.教室号", "学分制考试安排最后结果", "ccYsj.考试时间=#" & strKssj(i) & "#")
    k = DCount("教室号", "NotXfzResult", "考试时间=#" & strKssj(i) & "#")
    bKcs(i) = j + k
    iTotal = iTotal + bKcs(i)
    If iMaxKcs < bKcs(i) Then iMaxKcs = bKcs(i)
    rs.MoveNext
Loop

d = iTotal / iTeacher
iAvg = Int(2 * d + 1) '计算每个老师的平均场次,向上取整


Dim kch() As String
Dim jsh() As String

ReDim kch(iCount, iMaxKcs)
ReDim jsh(iCount, iMaxKcs)

For i = 1 To iCount
    str = "SELECT ccYsj.考试时间, jsYjkr.课程号, jsYjkr.教室号, jsYjkr.T1, jsYjkr.T2 " _
& "FROM (ccYsj INNER JOIN kcYbj ON ccYsj.场次标识 = kcYbj.场次标识) INNER JOIN jsYjkr ON (kcYbj.班级号 = jsYjkr.班级号) AND (kcYbj.课程号 = jsYjkr.课程号) " _
& "WHERE (((ccYsj.考试时间)=#" & strKssj(i) & "#))"
    Set rs = CurrentDb.OpenRecordset(str)
    j = 0
    Do While Not rs.EOF
        j = j + 1
        kch(i, j) = rs("课程号")
        jsh(i, j) = rs("教室号")
        rs.MoveNext
    Loop
    
    Set rs = CurrentDb.OpenRecordset("select * from NotXfzResult where 考试时间=#" & strKssj(i) & "#")
    Do While Not rs.EOF
        j = j + 1
        jsh(i, j) = rs("教室号")
        rs.MoveNext
    Loop
Next i

''''''''''''''''''''''''''''''''''''''''开始排监考老师
Dim iQH As Integer, iQT As Integer '定义正在使用的教师队列
Dim strTU As String '放置当前考试时间已经使用了的教师
Dim TQ() As Integer
ReDim TQ(2 * iTeacher)
For i = 1 To iTeacher
    TQ(i) = i
Next i

c = 0
k = 1
bBack = False
For i = 1 To iCount '按时间从第一场到最后一场排
    strTU = ";"
    For j = 1 To bKcs(i) '遍历某一时间的所有考场
        If i = 1 And j = 1 Then iQH = iQH + 1 '第一次进队列
        
        If Len(kch(i, j)) <> 0 Then '学分制的
            If Tcount(TQ(k)) = 0 Then iQT = iQT + 1 '进队
            CurrentDb.Execute ("update jsYjkr set T1='" & TQ(k) & "' where 课程号='" & kch(i, j) & "' and 教室号='" & jsh(i, j) & "'")
            strTU = strTU & TQ(k) & ";"
            Tcount(TQ(k)) = Tcount(TQ(k)) + 1
            If Tcount(TQ(k)) = iAvg Then iQH = iQH + 1  '出队
            'Debug.Print strTU
            
            k = k + 1
            If k > iTeacher Then
                For a = iTeacher To 1 Step -1 '已经排到了最后一个监考老师了，后退寻找。
                    If InStr(strTU, ";" & a & ";") = 0 And Tcount(TQ(a)) = iAvg Then
                        iQT = iQT + 1 '进队
                        TQ(k) = a
                        Exit For
                    End If
                Next a
            End If
            
            '排某个考场的第二个监考老师
            If DLookup("语音室", "kcYrs", "课程号='" & kch(i, j) & "'") Then
                b = b + 1 '记录听力考试的场次数,后面排听力考试的监考老师时要用到
            Else
                If Tcount(TQ(k)) = 0 Then iQT = iQT + 1 '进队
                CurrentDb.Execute ("update jsYjkr set T2='" & TQ(k) & "' where 课程号='" & kch(i, j) & "' and 教室号='" & jsh(i, j) & "'")
                strTU = strTU & TQ(k) & ";"
                Tcount(TQ(k)) = Tcount(TQ(k)) + 1
                If Tcount(TQ(k)) = iAvg Then iQH = iQH + 1 '出队
                'Debug.Print strTU
                
                k = k + 1
                If k > iTeacher Then
                    For a = iTeacher To 1 Step -1  '已经排到了最后一个监考老师了，后退寻找。
                        If InStr(strTU, ";" & a & ";") = 0 And Tcount(TQ(a)) = iAvg Then
                            iQT = iQT + 1 '进队
                            TQ(k) = a
                            Exit For
                        End If
                    Next a
                End If
            End If
        Else '非学分制的
            If Tcount(TQ(k)) = 0 Then iQT = iQT + 1 '进队
            'If Tcount(TQ(k)) < iAvg + 1 Then
            CurrentDb.Execute ("update NotXfzResult set T1='" & TQ(k) & "' where 考试时间=#" & strKssj(i) & "# and 教室号='" & jsh(i, j) & "'")
            strTU = strTU & TQ(k) & ";"
            Tcount(TQ(k)) = Tcount(TQ(k)) + 1
            If Tcount(TQ(k)) = iAvg Then iQH = iQH + 1 '出队
            'Debug.Print strTU
            
            k = k + 1
            If k > iTeacher Then
                For a = iTeacher To 1 Step -1 '已经排到了最后一个监考老师了，后退寻找。
                    If InStr(strTU, ";" & a & ";") = 0 And Tcount(TQ(a)) = iAvg Then
                        iQT = iQT + 1 '进队
                        TQ(k) = a
                        Exit For
                    End If
                Next a
            End If
            
            If Tcount(TQ(k)) = 0 Then iQT = iQT + 1 '进队
            'If Tcount(TQ(k)) < iAvg + 1 Then
            CurrentDb.Execute ("update NotXfzResult set T2='" & TQ(k) & "' where 考试时间=#" & strKssj(i) & "# and 教室号='" & jsh(i, j) & "'")
            strTU = strTU & TQ(k) & ";"
            Tcount(TQ(k)) = Tcount(TQ(k)) + 1
            If Tcount(TQ(k)) = iAvg Then iQH = iQH + 1 '出队
            'Debug.Print strTU
            
            k = k + 1
            If k > iTeacher Then
                For a = iTeacher To 1 Step -1 '已经排到了最后一个监考老师了，后退寻找。
                    If InStr(strTU, ";" & a & ";") = 0 And Tcount(TQ(a)) = iAvg Then
                        iQT = iQT + 1 '进队
                        TQ(k) = a
                        Exit For
                    End If
                Next a
            End If
        End If
    
        c = c + 1 '累计已经排完的考场数
    Next j
    

'    If (2 * d - Int(2 * d)) > 0 Then
'        d = 0.85
'    Else
'        d = 0.45
'    End If
    
    If Not bBack And c / iTotal > 0.48 Then  '这个参数表示排平均次数的老师的百分比,尽量大但是又要不出现兔子尾，有待优化修正【(2 * d-Int(2 * d))*(2 * d-Int(2 * d))】。
        c = 0 '只能进来一次
        bBack = True
        iAvg = iAvg - 1 '改变次数
        Do While Tcount(iQH) = iAvg  '将已经排了iAvg次的老师出队
            iQH = iQH + 1
        Loop
    End If
    
    If iQH <= iTeacher Then k = iQH  '找到队头
   
Next i

'''''''''''''''''''''''''排听力考试的监考老师'''''''''''''''''''''''''
iListen = DCount("*", "teacher", "听力否=True")
If iListen > 0 Then
    str = "SELECT teacher.单位, teacher.姓名 FROM teacher WHERE (听力否=True) GROUP BY teacher.单位, teacher.姓名 ORDER BY teacher.单位, teacher.姓名"
    Set rs = CurrentDb.OpenRecordset(str)
    j = rs.RecordCount
    rs.MoveLast
    rs.MoveFirst
    If iListen <> j Then
        MsgBox "监考老师表有重复的听力老师，请检查！"
        rs.Close
        Set rs = Nothing
        Exit Sub '退出
    End If
    
    ReDim Tdw(iListen)
    ReDim Tnm(iListen)
    ReDim Tcount(iListen)
    
    i = 0
    While Not rs.EOF    '对听力监考老师赋初值
        i = i + 1                                           '经
        j = Int(((iListen - i + 1) * Rnd) + 1)             '
        k = 0                                               '
        a = 0                                               '
        Do While k < j                                      '
            a = a + 1                                       '
            If Tnm(a) = "" Then k = k + 1                   '
        Loop                                                '典
        
        Tdw(a) = rs("单位")
        Tnm(a) = rs("姓名")
        Tcount(a) = 0
        CurrentDb.Execute ("update teacher set ID=" & iTeacher + a & " where 单位='" & Tdw(a) & "' and 姓名='" & Tnm(a) & "'")
        rs.MoveNext
    Wend
    iAvg = Int(b / iListen)
    
    ReDim TQ(2 * iListen)
    For i = 1 To iListen
        TQ(i) = i
    Next i
    
    iQH = 0: iQT = 0
    k = 1
    For i = 1 To iCount '按时间从第一场到最后一场排
        strTU = ";"
        For j = 1 To bKcs(i)
            If i = 1 And j = 1 Then iQH = iQH + 1 '第一次进队列
            If DLookup("语音室", "kcYrs", "课程号='" & kch(i, j) & "'") Then
                If TQ(k) = 0 Then
                    MsgBox "听力监考老师太少"
                    Exit Sub
                End If
                If Tcount(TQ(k)) = 0 Then iQT = iQT + 1 '进队
                CurrentDb.Execute ("update jsYjkr set T2='" & iTeacher + TQ(k) & "' where 课程号='" & kch(i, j) & "' and 教室号='" & jsh(i, j) & "'")
                strTU = strTU & TQ(k) & ";"
                Tcount(TQ(k)) = Tcount(TQ(k)) + 1
                If Tcount(TQ(k)) = iAvg Then iQH = iQH + 1 '出队
                
                k = k + 1
                If k > iListen Then
                    For a = iListen To 1 Step -1
                        If InStr(strTU, ";" & a & ";") = 0 And Tcount(TQ(a)) = iAvg Then
                            iQT = iQT + 1
                            TQ(k) = a
                            Exit For
                        End If
                    Next a
                End If
            End If
        Next j
        If iQH <= iListen Then k = iQH  '找到队头
    Next i
End If  'iListen

If MsgBox("监考人员顺利排完,是否开始生成Total表？", vbYesNo) <> vbYes Then
    Exit Sub
End If

If Not testTable("Total") Then
    CurrentDb.Execute ("SELECT 考试时间, 班级, 序号, 人数, 考试课程, 教室号, T1 as T INTO Total " _
    & "FROM NotXfzResult")
    CurrentDb.Execute ("delete from Total")
'    CurrentDb.Execute ("alter table Total add column 监考老师单位 text(20),监考老师 text(10)")
Else
    If CurrentData.AllTables("Total").IsLoaded Then DoCmd.Close acTable, "Total", acSaveNo
    CurrentDb.Execute ("delete from Total")
End If

str = "select 考试时间, 班级, 序号, 人数, 考试课程, 教室号 ,T1, T2 from NotXfzResult"
Set rs = CurrentDb.OpenRecordset(str)
Do While Not rs.EOF
    CurrentDb.Execute ("insert into Total (考试时间, 班级, 序号, 人数, 考试课程, 教室号,T) values ('" _
        & rs("考试时间") & "','" & rs("班级") & "'," & rs("序号") & "," & rs("人数") & ",'" & rs("考试课程") & "','" & rs("教室号") & "'," & rs("T1") & ")")
    CurrentDb.Execute ("insert into Total (考试时间, 班级, 序号, 人数, 考试课程, 教室号,T) values ('" _
        & rs("考试时间") & "','" & rs("班级") & "'," & rs("序号") & "," & rs("人数") & ",'" & rs("考试课程") & "','" & rs("教室号") & "'," & rs("T2") & ")")
    rs.MoveNext
Loop

str = "SELECT ccYsj.考试时间, 班级.班级名称, jsYjkr.序号, jsYjkr.班级人数, 课程.课程名称标识, jsYjkr.教室号, jsYjkr.T1, jsYjkr.T2 " _
& "FROM (ccYsj INNER JOIN kcYbj ON ccYsj.场次标识 = kcYbj.场次标识) INNER JOIN ((jsYjkr INNER JOIN 课程 ON jsYjkr.课程号 = 课程.课程号) INNER JOIN 班级 ON jsYjkr.班级号 = 班级.班级号) ON (kcYbj.班级号 = jsYjkr.班级号) AND (kcYbj.课程号 = jsYjkr.课程号)"
Set rs = CurrentDb.OpenRecordset(str)
Do While Not rs.EOF
    CurrentDb.Execute ("insert into Total (考试时间, 班级, 序号, 人数, 考试课程, 教室号,T) values ('" _
        & rs("考试时间") & "','" & rs("班级名称") & "'," & rs("序号") & "," & rs("班级人数") & ",'" & rs("课程名称标识") & "','" & rs("教室号") & "'," & rs("T1") & ")")
    CurrentDb.Execute ("insert into Total (考试时间, 班级, 序号, 人数, 考试课程, 教室号,T) values ('" _
        & rs("考试时间") & "','" & rs("班级名称") & "'," & rs("序号") & "," & rs("班级人数") & ",'" & rs("课程名称标识") & "','" & rs("教室号") & "'," & rs("T2") & ")")
    rs.MoveNext
Loop

Exit Sub
退出_Click_error:
MsgBox "参数太小了，请调大。"
End Sub
