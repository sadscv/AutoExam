Option Compare Database
Option Explicit
Option Base 1

Private Sub 确定_Click()
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
Dim jOk() As Byte '教室已排
Dim jType() As String '教室类型 1普通，2多媒体，3语音室


ReDim jNo(iClassroom)
ReDim jSize(iClassroom)
ReDim jOk(iClassroom)
ReDim jType(iClassroom)

str = "select 教室号,教室定员,教室类型号 from 教室 where 考试否 and 使用优先级>0 order by 使用优先级" '20130611优先级
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
str = "SELECT kcYrs.场次标识, xsYxk.课程号, xsYxk.班级号, Count(xsYxk.学号) AS 班级人数 INTO kcYbj " _
& "FROM kcYrs INNER JOIN xsYxk ON kcYrs.课程号 = xsYxk.课程号 " _
& "GROUP BY kcYrs.场次标识, xsYxk.课程号, xsYxk.班级号 " _
& "ORDER BY kcYrs.场次标识"

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

iCount = DMax("场次标识", "kcYbj")

For i = 1 To iCount
    For j = 1 To iClassroom
        jOk(j) = 0
    Next j
    
    str = "select * from kcYbj where 场次标识=" & i & " order by 课程号" '需要进行优化 desc
    Set rs = CurrentDb.OpenRecordset(str)
    rs.MoveLast
    rs.MoveFirst
    Dim iA As Integer
    Dim iB As Integer
    Dim iC As Integer
    Dim m As Integer, n As Integer, p As Integer
    Dim bok As Boolean '这门课程是否已排定
    
    Do While Not rs.EOF
        '先排要使用语音室的课程
        bok = False
        If DLookup("语音室", "kcYrs", "课程号='" & Trim(rs("课程号")) & "'") Then
            If rs("班级人数") > 60 Then
                iA = rs("班级人数") / 2
                iB = rs("班级人数") - iA
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',2)")
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iB, 2)
                For j = 1 To iClassroom - 1
                    If iA <= jSize(j) And jOk(j) = 0 And Not bok And jType(j) = "3" Then
                        For k = j + 1 To iClassroom
                            If iB <= jSize(k) And jOk(k) = 0 And Left(jNo(k), 1) = Left(jNo(j), 1) And jType(k) = "3" Then
                                CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(j) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=1")
                                jOk(j) = 1
                                CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(k) & "', 班级人数=" & iB & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=2")
                                jOk(k) = 1
                                bok = True
                                Exit For '这个班级排定
                            End If
                        Next k
                    End If
                    
                    If bok Then Exit For
                Next j
            Else
                For j = 1 To iClassroom
                    If rs("班级人数") <= jSize(j) And jOk(j) = 0 And jType(j) = "3" Then
                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(j) & "' where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "'")
                        jOk(j) = 1
                        Exit For '这个班级排定
                    End If
                Next j
                If j = iClassroom + 1 Then
                    MsgBox rs("课程号") & "未排好"
                End If
            End If
        End If
        
        rs.MoveNext
    Loop
    rs.MoveFirst
    
    Do While Not rs.EOF
        '再排不使用语音室的课程
        bok = False
        If Not DLookup("语音室", "kcYrs", "课程号='" & Trim(rs("课程号")) & "'") Then
            '如果人数大于240，拆成5个班
            If rs("班级人数") > 240 Then
                iA = rs("班级人数") / 5
                iB = rs("班级人数") - iA * 4
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',2)")
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',3)")
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',4)")
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',5)")
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iA, 2)
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iA, 3)
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iA, 4)
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iB, 5)
                For j = 1 To iClassroom - 4
                    If iA <= jSize(j) And jOk(j) = 0 And Not bok Then
                        For k = j + 1 To iClassroom - 3
                            If iA <= jSize(k) And jOk(k) = 0 And Left(jNo(k), 1) = Left(jNo(j), 1) And Not bok Then
                                For m = k + 1 To iClassroom - 2
                                    If iA <= jSize(m) And jOk(m) = 0 And Left(jNo(m), 1) = Left(jNo(j), 1) And Not bok Then
                                        For n = m + 1 To iClassroom - 1
                                            If iA < jSize(n) And jOk(n) = 0 And Left(jNo(n), 1) = Left(jNo(j), 1) And Not bok Then
                                                For p = n + 1 To iClassroom
                                                    If iB <= jSize(p) And jOk(p) = 0 And Left(jNo(p), 1) = Left(jNo(j), 1) Then
                                                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(j) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=1")
                                                        jOk(j) = 1
                                                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(k) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=2")
                                                        jOk(k) = 1
                                                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(m) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=3")
                                                        jOk(m) = 1
                                                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(n) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=4")
                                                        jOk(n) = 1
                                                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(p) & "', 班级人数=" & iB & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=5")
                                                        jOk(p) = 1
                                                        bok = True
                                                        Exit For '这个班级排定
                                                    End If
                                                Next p
                                            End If
                                            
                                            If bok Then Exit For
                                        Next n
                                    End If
                                    
                                    If bok Then Exit For
                                Next m
                            End If
                            
                            If bok Then Exit For
                        Next k
                    End If
                    
                    If bok Then Exit For
                Next j
            '如果人数大于180，拆成4个班
            ElseIf rs("班级人数") > 180 Then
                iA = rs("班级人数") / 4
                iB = rs("班级人数") - iA * 3
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',2)")
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',3)")
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',4)")
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iA, 2)
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iA, 3)
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iB, 4)
                For j = 1 To iClassroom - 3
                    If iA <= jSize(j) And jOk(j) = 0 And Not bok Then
                        For k = j + 1 To iClassroom - 2
                            If iA <= jSize(k) And jOk(k) = 0 And Left(jNo(k), 1) = Left(jNo(j), 1) And Not bok Then
                                For m = k + 1 To iClassroom - 1
                                    If iA <= jSize(m) And jOk(m) = 0 And Left(jNo(m), 1) = Left(jNo(j), 1) And Not bok Then
                                        For n = m + 1 To iClassroom
                                            If iB <= jSize(n) And jOk(n) = 0 And Left(jNo(n), 1) = Left(jNo(j), 1) Then
                                                CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(j) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=1")
                                                jOk(j) = 1
                                                CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(k) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=2")
                                                jOk(k) = 1
                                                CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(m) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=3")
                                                jOk(m) = 1
                                                CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(n) & "', 班级人数=" & iB & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=4")
                                                jOk(n) = 1
                                                bok = True
                                                Exit For '这个班级排定
                                            End If
                                        Next n
                                    End If
                                    
                                    If bok Then Exit For
                                Next m
                            End If
                            
                            If bok Then Exit For
                        Next k
                    End If
                    
                    If bok Then Exit For
                Next j
            '如果人数大于120，拆成3个班
            ElseIf rs("班级人数") > 120 Then
                iA = rs("班级人数") / 3
                iB = rs("班级人数") - iA * 2
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',2)")
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',3)")
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iA, 2)
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iB, 3)
                For j = 1 To iClassroom - 2
                    If iA <= jSize(j) And jOk(j) = 0 And Not bok Then
                        For k = j + 1 To iClassroom - 1
                            If iA <= jSize(k) And jOk(k) = 0 And Left(jNo(k), 1) = Left(jNo(j), 1) And Not bok Then
                                For m = k + 1 To iClassroom
                                    If iB <= jSize(m) And jOk(m) = 0 And Left(jNo(m), 1) = Left(jNo(j), 1) Then
                                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(j) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=1")
                                        jOk(j) = 1
                                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(k) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=2")
                                        jOk(k) = 1
                                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(m) & "', 班级人数=" & iB & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=3")
                                        jOk(m) = 1
                                        bok = True
                                        Exit For '这个班级排定
                                    End If
                                Next m
                            End If
                            
                            If bok Then Exit For
                        Next k
                    End If
                    
                    If bok Then Exit For
                Next j
            '如果人数大于60，拆成2个班
            ElseIf rs("班级人数") > 60 Then
                iA = rs("班级人数") / 2
                iB = rs("班级人数") - iA
                CurrentDb.Execute ("insert into jsYjkr (场次标识, 课程号,班级号,序号) values ('" & rs("场次标识") & "','" & rs("课程号") & "','" & rs("班级号") & "',2)")
                bok = bok And Update_xsYxk(Trim(rs("课程号")), Trim(rs("班级号")), iB, 2)
                For j = 1 To iClassroom - 1
                    If iA <= jSize(j) And jOk(j) = 0 And Not bok Then
                        For k = j + 1 To iClassroom
                            If iB <= jSize(k) And jOk(k) = 0 And Left(jNo(k), 1) = Left(jNo(j), 1) Then
                                CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(j) & "', 班级人数=" & iA & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=1")
                                jOk(j) = 1
                                CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(k) & "', 班级人数=" & iB & " where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "' and 序号=2")
                                jOk(k) = 1
                                bok = True
                                Exit For '这个班级排定
                            End If
                        Next k
                    End If
                    
                    If bok Then Exit For
                Next j
            Else
                For j = 1 To iClassroom
                    If rs("班级人数") <= jSize(j) And jOk(j) = 0 Then
                        CurrentDb.Execute ("update jsYjkr set 教室号='" & jNo(j) & "' where 课程号='" & Trim(rs("课程号")) & "' and 班级号='" & Trim(rs("班级号")) & "'")
                        jOk(j) = 1
                        Exit For '这个班级排定
                    End If
                Next j
            End If
        End If
        
        If j = iClassroom + 1 Then '这个班级人数太多,只好分多个教室来排,但注意要放在同一楼，最好是同一层。
            MsgBox "第" & i & "场的" & rs("课程号") & "没法排"
        End If
        
        rs.MoveNext
    Loop
Next i
MsgBox "教室已排完"
Me.退出.SetFocus
Me.确定.Enabled = False
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

Private Sub 退出_Click()
    Dim str As String
    Dim ifRecreate As Boolean
    Dim rs As DAO.Recordset
    
    str = "SELECT kcYbj.场次标识, '' AS 备注 INTO ccYsj " _
            & "FROM kcYbj where 场次标识>0 GROUP BY kcYbj.场次标识, '', ''"
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

    '做2010-9-1排考时新加------------------------------------------------20130611好像确实有必要
    If Not testTable("NotXfz") Then
        MsgBox ("非学分制的考试安排表不存在，请导入！")
        Exit Sub
    Else
        Dim iLs As Integer
        iLs = DMax("场次标识", "ccYsj")
        str = "select distinct 考试时间 from NotXfz"
        Set rs = CurrentDb.OpenRecordset(str)
        rs.MoveLast
        rs.MoveFirst
        Do While Not rs.EOF
            If Nz(DLookup("考试时间", "ccYsj", "考试时间" = rs("考试时间")), 0) = 0 Then
                iLs = iLs + 1
                CurrentDb.Execute ("insert into ccYsj (场次标识, 考试时间) values (" & iLs & ",#" & rs("考试时间") & "#)")
            End If
            rs.MoveNext
        Loop
    End If '------------------------------------------------------------------------------------------
    
    '检查是否存生成了“非学分制场次与教室总表”AllccYjs
    If Not testTable("AllccYjs") Then
        ifRecreate = True
    ElseIf MsgBox("非学分制表场次与教室总表已经存在，是否重建？", vbYesNo, "询问") = vbYes Then
        ifRecreate = True
    End If
    
    If ifRecreate Then
        If testTable("AllccYjs") Then
            If CurrentData.AllTables("AllccYjs").IsLoaded Then DoCmd.Close acTable, "AllccYjs", acSaveNo
            CurrentDb.Execute ("drop table AllccYjs")
        End If
    
        str = "SELECT ccYsj.场次标识, 教室.教室号 INTO AllccYjs FROM ccYsj, 教室 WHERE (((教室.考试否) = True) and ((教室.使用优先级)>0)) ORDER BY ccYsj.场次标识, 教室.教室号"
        CurrentDb.Execute (str)
        CurrentDb.Execute ("alter table AllccYjs add column 已用否 bit")
        
        '把已经使用了的教室从“非学分制场次与教室总表”中删除
        'Dim rs As DAO.Recordset
        str = "SELECT kcYbj.场次标识, jsYjkr.教室号 " _
            & "FROM kcYbj INNER JOIN jsYjkr ON (kcYbj.班级号 = jsYjkr.班级号) AND (kcYbj.课程号 = jsYjkr.课程号)"
        Set rs = CurrentDb.OpenRecordset(str)
        While Not rs.EOF
            CurrentDb.Execute ("delete from AllccYjs where 场次标识=" & Trim(rs("场次标识")) & " and 教室号='" & Trim(rs("教室号")) & "'")
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
    End If
    
    If DLookup("场次标识", "ccYsj", "考试时间 is null") Then
        MsgBox "请注意，还有考试时间没有录入！"
    End If
    
    DoCmd.Close , , acSaveYes
    DoCmd.OpenForm "考试时间录入", acNormal
End Sub

