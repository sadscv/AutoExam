Option Compare Database
Option Explicit
Dim iRS As Integer

Private Sub Form_Load()
'检查参数表，如果不存在就新建并且初始化
If Not testTable("Para") Then
    Dim str As String
    
    str = "create table Para(ksxq datetime,A integer, B byte)"
    CurrentDb.Execute (str)
    
    str = "insert into Para values (#3/1/2004#,0,0)"
    CurrentDb.Execute (str)
End If
End Sub

Private Sub 考试学期_AfterUpdate()
'更新参数表的考试学期
If Not testTable("Para") Then
    MsgBox ("表Para不存在")
    Exit Sub
End If

Dim str As String

str = "update Para set ksxq=#" & Me.考试学期 & "#"
CurrentDb.Execute (str)
End Sub

Private Sub 确定_Click()
Dim str As String
Dim rs As DAO.Recordset
Dim i As Integer


If Nz(Me.考试学期) = "" Then
    MsgBox ("请选择考试学期")
    Exit Sub
End If


' 删除课程与人数表
If testTable("kcYrs") Then
    '判断该表是否打开，若打开则强行关闭
    If CurrentData.AllTables("kcYrs").IsLoaded Then DoCmd.Close acTable, "kcYrs", acSaveNo
    CurrentDb.Execute ("drop table kcYrs")
End If


' 课程号	选课人数	场次标识	不考试	语音室
' 266032    	11	6	0	0
' 266115    	15	16	0	0
' 266099    	15	8	0	0
' 062313    	15	22	0	0
' 266077    	16	6	0	0
str = "SELECT 学生与选课.课程号, Count(学生与选课.学号) AS 选课人数, 0 AS 场次标识 INTO kcYrs " _
& "FROM (班级 INNER JOIN 学生 ON 班级.班级号 = 学生.班级号) INNER JOIN 学生与选课 ON 学生.学号 = 学生与选课.学号 " _
& "WHERE (((学生与选课.开课时间) =#" & Me.考试学期 & "#) And ((学生与选课.选课状态) = 0)) " _
& "GROUP BY 学生与选课.课程号 ORDER BY Count(学生与选课.学号)"
'Debug.Print str
CurrentDb.Execute (str)

str = "alter table kcYrs add column 不考试 bit"
CurrentDb.Execute (str)
str = "alter table kcYrs add column 语音室 bit"
CurrentDb.Execute (str)

'str = "update kcYrs set 不考试=true where 课程号 like '00*'"
'CurrentDb.Execute (str)

'加上改变kcYrs表的“不考试”字段的显示控件为CheckBox的语句
'With CurrentDb.TableDefs("kcYrs").Fields(不考试).Properties
'    .VisibleValue = True
'End With

DoCmd.Close , , acSaveYes
DoCmd.OpenForm "排场次"
End Sub

Private Sub 退出_Click()
If MsgBox("你确定要退出吗？", vbYesNo, "提示") = vbYes Then
    DoCmd.Close
End If
End Sub
