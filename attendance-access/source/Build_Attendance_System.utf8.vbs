'==========================================================================
' نظام الحضور والانصراف والأجور الإضافية - Microsoft Access
' Builder script: creates Attendance_System.accdb next to this file with
' tables, relationships, sample data (10 cases), queries, forms, 10 reports.
' Requirements: Windows + Microsoft Access 2010 or later (any bitness).
' Usage: double-click Build_Attendance_System.vbs
'==========================================================================
Option Explicit

Const acForm = 2, acReport = 3, acSaveYes = 1
Const acLabel = 100, acRectangle = 101, acCommandButton = 104
Const acCheckBox = 106, acTextBox = 109, acComboBox = 111
Const acDetail = 0, acPageHeader = 3, acPageFooter = 4
Const dbText = 10, dbOpenDynaset = 2, dbFailOnError = 128
Const CM = 567
Const FONT_NAME = "Arial"

Dim fso, app, db, dbPath, gLog, ans, lf
Dim P_FROM, P_TO, P_EMP, P_DEPT, PARAM_DATES
P_FROM = "[Forms]![frmMain]![txtFrom]"
P_TO = "[Forms]![frmMain]![txtTo]"
P_EMP = "[Forms]![frmMain]![cboEmp]"
P_DEPT = "[Forms]![frmMain]![cboDept]"
PARAM_DATES = "PARAMETERS " & P_FROM & " DateTime, " & P_TO & " DateTime; "
gLog = ""

Set fso = CreateObject("Scripting.FileSystemObject")
dbPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Attendance_System.accdb"

If fso.FileExists(dbPath) Then
    ans = MsgBox("الملف موجود بالفعل:" & vbCrLf & dbPath & vbCrLf & vbCrLf & _
                 "هل تريد حذفه وإعادة بنائه من جديد؟", vbYesNo + vbQuestion, "نظام الحضور والانصراف")
    If ans <> vbYes Then WScript.Quit
    fso.DeleteFile dbPath, True
End If

MsgBox "سيتم الآن بناء قاعدة البيانات (قد يستغرق ذلك دقيقة أو دقيقتين)." & vbCrLf & _
       "اضغط OK وانتظر رسالة الانتهاء.", vbInformation, "نظام الحضور والانصراف"

Set app = CreateObject("Access.Application")
app.Visible = False
app.NewCurrentDatabase dbPath, 12
Set db = app.CurrentDb

CreateTables
FillSettingsAndLookups
FillCalendarHolidaysAndSampleData
CreateQueries
CreateForms
CreateReports
FinalizeDb

app.CloseCurrentDatabase
app.Quit
Set app = Nothing

If gLog <> "" Then
    Set lf = fso.CreateTextFile(fso.GetParentFolderName(WScript.ScriptFullName) & "\build_log.txt", True, True)
    lf.Write gLog
    lf.Close
    MsgBox "تم البناء مع بعض التحذيرات (راجع ملف build_log.txt)." & vbCrLf & dbPath, vbExclamation, "نظام الحضور والانصراف"
Else
    MsgBox "تم بناء البرنامج بنجاح:" & vbCrLf & dbPath & vbCrLf & vbCrLf & _
           "عند فتح الملف اضغط (تمكين المحتوى / Enable Content).", vbInformation, "نظام الحضور والانصراف"
End If
CreateObject("WScript.Shell").Run """" & dbPath & """"

'==========================================================================
' Helpers
'==========================================================================
Sub LogErr(ctx)
    If Err.Number <> 0 Then
        gLog = gLog & ctx & " : " & Err.Number & " - " & Err.Description & vbCrLf
        Err.Clear
    End If
End Sub

Sub Exec(sql, ctx)
    On Error Resume Next
    db.Execute sql, dbFailOnError
    LogErr ctx
End Sub

Sub MakeQ(qName, sql)
    On Error Resume Next
    db.CreateQueryDef qName, sql
    LogErr "Query " & qName
End Sub

Sub SetFldProp(tbl, fld, pName, pType, pVal)
    On Error Resume Next
    Dim f
    Set f = db.TableDefs(tbl).Fields(fld)
    f.Properties(pName) = pVal
    If Err.Number <> 0 Then
        Err.Clear
        f.Properties.Append f.CreateProperty(pName, pType, pVal)
    End If
    LogErr "FieldProp " & tbl & "." & fld & "." & pName
End Sub

Sub SetDbProp(pName, pType, pVal)
    On Error Resume Next
    db.Properties(pName) = pVal
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty(pName, pType, pVal)
    End If
    LogErr "DbProp " & pName
End Sub

' captions: Array(field, caption, field, caption, ...)
Sub Captions(tbl, arr)
    Dim i
    For i = 0 To UBound(arr) Step 2
        SetFldProp tbl, arr(i), "Caption", dbText, arr(i + 1)
    Next
End Sub

Function T(h, m)
    T = TimeSerial(h, m, 0)
End Function

Function DayNameExpr(f)
    DayNameExpr = "Choose(Weekday(" & f & "),""الأحد"",""الاثنين"",""الثلاثاء"",""الأربعاء"",""الخميس"",""الجمعة"",""السبت"")"
End Function

Function NZ0(expr)
    NZ0 = "IIf(" & expr & " Is Null,0," & expr & ")"
End Function

Function SafeName(s)
    Dim r
    r = Replace(s, ".", "")
    r = Replace(r, "!", "")
    r = Replace(r, "[", "")
    r = Replace(r, "]", "")
    r = Replace(r, "`", "")
    SafeName = r
End Function

'==========================================================================
' 1) Tables & relationships
'==========================================================================
Sub CreateTables()
    On Error Resume Next
    Exec "CREATE TABLE tblSettings (ID LONG CONSTRAINT pkSettings PRIMARY KEY, CompanyName TEXT(100), CurrencyName TEXT(20), " & _
         "GraceMinutes LONG, OTStartAfterMin LONG, OTMinMinutes LONG, NightStart DATETIME, " & _
         "OTRateDay DOUBLE, OTRateNight DOUBLE, OTRateOff DOUBLE, LunchAmount CURRENCY, DinnerAmount CURRENCY, " & _
         "WorkDaysPerMonth DOUBLE, WeekendDay1 LONG, WeekendDay2 LONG, LateDeductOn YESNO, AbsenceDeductDays DOUBLE)", "tblSettings"
    Exec "CREATE TABLE tblDepartments (DeptID COUNTER CONSTRAINT pkDept PRIMARY KEY, DeptName TEXT(100) NOT NULL)", "tblDepartments"
    Exec "CREATE TABLE tblShifts (ShiftID COUNTER CONSTRAINT pkShift PRIMARY KEY, ShiftName TEXT(100) NOT NULL, " & _
         "StartTime DATETIME NOT NULL, EndTime DATETIME NOT NULL, LunchAfter DATETIME, DinnerAfter DATETIME)", "tblShifts"
    Exec "CREATE TABLE tblEmployees (EmpID LONG CONSTRAINT pkEmp PRIMARY KEY, EmpName TEXT(150) NOT NULL, DeptID LONG, " & _
         "JobTitle TEXT(100), ShiftID LONG NOT NULL, HireDate DATETIME, BasicSalary CURRENCY, AnnualLeaveDays LONG, " & _
         "Phone TEXT(30), Active YESNO, Notes TEXT(255), " & _
         "CONSTRAINT fkEmpDept FOREIGN KEY (DeptID) REFERENCES tblDepartments (DeptID), " & _
         "CONSTRAINT fkEmpShift FOREIGN KEY (ShiftID) REFERENCES tblShifts (ShiftID))", "tblEmployees"
    Exec "CREATE TABLE tblLeaveTypes (LeaveTypeID COUNTER CONSTRAINT pkLT PRIMARY KEY, LeaveTypeName TEXT(100) NOT NULL, " & _
         "IsPaid YESNO, DefaultBalance LONG, UseEmpBalance YESNO)", "tblLeaveTypes"
    Exec "CREATE TABLE tblHolidays (HolidayDate DATETIME CONSTRAINT pkHol PRIMARY KEY, HolidayName TEXT(150))", "tblHolidays"
    Exec "CREATE TABLE tblCalendar (CalDate DATETIME CONSTRAINT pkCal PRIMARY KEY)", "tblCalendar"
    Exec "CREATE TABLE tblAttendance (AttID COUNTER CONSTRAINT pkAtt PRIMARY KEY, EmpID LONG NOT NULL, WorkDate DATETIME NOT NULL, " & _
         "TimeIn DATETIME, TimeOut DATETIME, Notes TEXT(255), " & _
         "CONSTRAINT fkAttEmp FOREIGN KEY (EmpID) REFERENCES tblEmployees (EmpID))", "tblAttendance"
    Exec "CREATE UNIQUE INDEX uxAttEmpDate ON tblAttendance (EmpID, WorkDate)", "uxAttEmpDate"
    Exec "CREATE INDEX ixAttDate ON tblAttendance (WorkDate)", "ixAttDate"
    Exec "CREATE TABLE tblLeaves (LeaveID COUNTER CONSTRAINT pkLeave PRIMARY KEY, EmpID LONG NOT NULL, LeaveTypeID LONG NOT NULL, " & _
         "FromDate DATETIME NOT NULL, ToDate DATETIME NOT NULL, Notes TEXT(255), " & _
         "CONSTRAINT fkLvEmp FOREIGN KEY (EmpID) REFERENCES tblEmployees (EmpID), " & _
         "CONSTRAINT fkLvType FOREIGN KEY (LeaveTypeID) REFERENCES tblLeaveTypes (LeaveTypeID))", "tblLeaves"
    db.TableDefs.Refresh

    ' table-level rules
    db.TableDefs("tblSettings").ValidationRule = "[ID]=1"
    db.TableDefs("tblSettings").ValidationText = "جدول الإعدادات يحتوي على سجل واحد فقط (ID = 1)"
    db.TableDefs("tblLeaves").ValidationRule = "[ToDate]>=[FromDate]"
    db.TableDefs("tblLeaves").ValidationText = "تاريخ نهاية الإجازة يجب أن يكون بعد أو يساوي تاريخ البداية"
    LogErr "Validation rules"
    db.TableDefs("tblEmployees").Fields("Active").DefaultValue = "True"
    db.TableDefs("tblEmployees").Fields("AnnualLeaveDays").DefaultValue = "21"
    db.TableDefs("tblLeaveTypes").Fields("IsPaid").DefaultValue = "True"
    db.TableDefs("tblLeaveTypes").Fields("DefaultBalance").DefaultValue = "0"
    LogErr "Default values"

    ' display formats
    SetFldProp "tblSettings", "NightStart", "Format", dbText, "hh:nn"
    SetFldProp "tblShifts", "StartTime", "Format", dbText, "hh:nn"
    SetFldProp "tblShifts", "EndTime", "Format", dbText, "hh:nn"
    SetFldProp "tblShifts", "LunchAfter", "Format", dbText, "hh:nn"
    SetFldProp "tblShifts", "DinnerAfter", "Format", dbText, "hh:nn"
    SetFldProp "tblAttendance", "TimeIn", "Format", dbText, "hh:nn"
    SetFldProp "tblAttendance", "TimeOut", "Format", dbText, "hh:nn"
    SetFldProp "tblAttendance", "WorkDate", "Format", dbText, "dd/mm/yyyy"
    SetFldProp "tblLeaves", "FromDate", "Format", dbText, "dd/mm/yyyy"
    SetFldProp "tblLeaves", "ToDate", "Format", dbText, "dd/mm/yyyy"
    SetFldProp "tblHolidays", "HolidayDate", "Format", dbText, "dd/mm/yyyy"
    SetFldProp "tblEmployees", "HireDate", "Format", dbText, "dd/mm/yyyy"
    SetFldProp "tblCalendar", "CalDate", "Format", dbText, "dd/mm/yyyy"

    ' Arabic captions
    Captions "tblSettings", Array("CompanyName", "اسم الشركة", "CurrencyName", "العملة", "GraceMinutes", "سماحية التأخير (دقيقة)", _
        "OTStartAfterMin", "بدء الإضافي بعد نهاية الوردية (دقيقة)", "OTMinMinutes", "الحد الأدنى للإضافي (دقيقة)", _
        "NightStart", "بداية الإضافي الليلي", "OTRateDay", "معامل الإضافي النهاري", "OTRateNight", "معامل الإضافي الليلي", _
        "OTRateOff", "معامل العمل في الراحات والعطلات", "LunchAmount", "قيمة بدل الغداء", "DinnerAmount", "قيمة بدل العشاء", _
        "WorkDaysPerMonth", "أيام الشهر لحساب الأجر اليومي", "WeekendDay1", "يوم الراحة 1", "WeekendDay2", "يوم الراحة 2", _
        "LateDeductOn", "تفعيل خصم التأخير", "AbsenceDeductDays", "أيام الخصم عن كل يوم غياب")
    Captions "tblDepartments", Array("DeptID", "كود الإدارة", "DeptName", "اسم الإدارة")
    Captions "tblShifts", Array("ShiftID", "كود الوردية", "ShiftName", "اسم الوردية", "StartTime", "بداية الدوام", "EndTime", "نهاية الدوام", _
        "LunchAfter", "بدل الغداء عند الانصراف بعد", "DinnerAfter", "بدل العشاء عند الانصراف بعد")
    Captions "tblEmployees", Array("EmpID", "كود الموظف", "EmpName", "اسم الموظف", "DeptID", "الإدارة", "JobTitle", "الوظيفة", _
        "ShiftID", "الوردية", "HireDate", "تاريخ التعيين", "BasicSalary", "الأجر الأساسي الشهري", "AnnualLeaveDays", "رصيد الاعتيادي السنوي", _
        "Phone", "الهاتف", "Active", "على رأس العمل", "Notes", "ملاحظات")
    Captions "tblLeaveTypes", Array("LeaveTypeID", "كود النوع", "LeaveTypeName", "نوع الإجازة", "IsPaid", "مدفوعة الأجر", _
        "DefaultBalance", "الرصيد السنوي", "UseEmpBalance", "يُستخدم رصيد الموظف")
    Captions "tblHolidays", Array("HolidayDate", "تاريخ العطلة", "HolidayName", "اسم العطلة")
    Captions "tblCalendar", Array("CalDate", "التاريخ")
    Captions "tblAttendance", Array("AttID", "رقم الحركة", "EmpID", "الموظف", "WorkDate", "التاريخ", "TimeIn", "وقت الحضور", _
        "TimeOut", "وقت الانصراف", "Notes", "ملاحظات")
    Captions "tblLeaves", Array("LeaveID", "رقم الإجازة", "EmpID", "الموظف", "LeaveTypeID", "نوع الإجازة", "FromDate", "من تاريخ", _
        "ToDate", "إلى تاريخ", "Notes", "ملاحظات")
End Sub

'==========================================================================
' 2) Settings & lookup data
'==========================================================================
Sub FillSettingsAndLookups()
    On Error Resume Next
    Dim rs, i, arr

    Set rs = db.OpenRecordset("tblSettings", dbOpenDynaset)
    rs.AddNew
    rs("ID") = 1
    rs("CompanyName") = "الشركة النموذجية للتجارة والخدمات"
    rs("CurrencyName") = "جنيه"
    rs("GraceMinutes") = 15
    rs("OTStartAfterMin") = 30
    rs("OTMinMinutes") = 30
    rs("NightStart") = T(19, 0)
    rs("OTRateDay") = 1.35
    rs("OTRateNight") = 1.7
    rs("OTRateOff") = 2
    rs("LunchAmount") = 50
    rs("DinnerAmount") = 75
    rs("WorkDaysPerMonth") = 30
    rs("WeekendDay1") = 6
    rs("WeekendDay2") = 7
    rs("LateDeductOn") = False
    rs("AbsenceDeductDays") = 1
    rs.Update
    rs.Close
    LogErr "Settings data"

    Set rs = db.OpenRecordset("tblDepartments", dbOpenDynaset)
    arr = Array("الموارد البشرية", "الشؤون المالية", "تكنولوجيا المعلومات", "العمليات", "خدمة العملاء")
    For i = 0 To UBound(arr)
        rs.AddNew
        rs("DeptName") = arr(i)
        rs.Update
    Next
    rs.Close
    LogErr "Departments data"

    Set rs = db.OpenRecordset("tblShifts", dbOpenDynaset)
    rs.AddNew
    rs("ShiftName") = "الوردية الصباحية"
    rs("StartTime") = T(9, 0)
    rs("EndTime") = T(17, 0)
    rs("LunchAfter") = T(18, 0)
    rs("DinnerAfter") = T(21, 0)
    rs.Update
    rs.AddNew
    rs("ShiftName") = "الوردية المسائية"
    rs("StartTime") = T(14, 0)
    rs("EndTime") = T(22, 0)
    rs("LunchAfter") = T(23, 0)
    rs("DinnerAfter") = T(0, 30)
    rs.Update
    rs.Close
    LogErr "Shifts data"

    ' name, paid, default balance, use employee balance
    arr = Array( _
        Array("إجازة اعتيادية", True, 0, True), _
        Array("إجازة عارضة", True, 6, False), _
        Array("إجازة مرضية", True, 0, False), _
        Array("إجازة بدون أجر", False, 0, False), _
        Array("مأمورية عمل", True, 0, False), _
        Array("إجازة وضع / رعاية طفل", True, 0, False))
    Set rs = db.OpenRecordset("tblLeaveTypes", dbOpenDynaset)
    For i = 0 To UBound(arr)
        rs.AddNew
        rs("LeaveTypeName") = arr(i)(0)
        rs("IsPaid") = arr(i)(1)
        rs("DefaultBalance") = arr(i)(2)
        rs("UseEmpBalance") = arr(i)(3)
        rs.Update
    Next
    rs.Close
    LogErr "Leave types data"

    ' EmpID, name, dept, job, shift, salary, annual leave, phone, note (the 10 cases)
    arr = Array( _
        Array(1001, "أحمد محمد علي", 3, "مطور نظم", 1, 15000, 21, "01000000001", "الحالة 1: موظف منضبط + مأمورية عمل يوم واحد"), _
        Array(1002, "محمد حسن إبراهيم", 4, "مشرف عمليات", 1, 12000, 21, "01000000002", "الحالة 2: تأخير متكرر وانصراف مبكر + يوم بدون أجر"), _
        Array(1003, "سارة أحمد محمود", 2, "محاسب أول", 1, 14000, 21, "01000000003", "الحالة 3: ساعات إضافية نهارية + بدل غداء"), _
        Array(1004, "محمود عبد الله سالم", 4, "أخصائي عمليات", 1, 11000, 21, "01000000004", "الحالة 4: ساعات إضافية ليلية + بدل غداء وعشاء"), _
        Array(1005, "منى سمير فؤاد", 1, "أخصائي موارد بشرية", 1, 13000, 30, "01000000005", "الحالة 5: إجازة اعتيادية 5 أيام"), _
        Array(1006, "خالد يوسف عمر", 5, "ممثل خدمة عملاء", 1, 8000, 21, "01000000006", "الحالة 6: غياب 3 أيام بدون إذن"), _
        Array(1007, "ياسمين طارق حسين", 2, "محاسب", 1, 10000, 21, "01000000007", "الحالة 7: إجازة مرضية + إجازة عارضة"), _
        Array(1008, "عمر شريف كامل", 3, "مهندس شبكات", 1, 16000, 21, "01000000008", "الحالة 8: عمل في يوم الراحة وفي عطلة رسمية"), _
        Array(1009, "هبة مصطفى رضا", 5, "مشرف خدمة عملاء", 1, 9500, 21, "01000000009", "الحالة 9: انصراف مبكر + بصمة انصراف ناقصة"), _
        Array(1010, "يوسف إبراهيم نصر", 4, "مراقب تشغيل", 2, 10500, 21, "01000000010", "الحالة 10: وردية مسائية + إضافي يتجاوز منتصف الليل"))
    Set rs = db.OpenRecordset("tblEmployees", dbOpenDynaset)
    For i = 0 To UBound(arr)
        rs.AddNew
        rs("EmpID") = arr(i)(0)
        rs("EmpName") = arr(i)(1)
        rs("DeptID") = arr(i)(2)
        rs("JobTitle") = arr(i)(3)
        rs("ShiftID") = arr(i)(4)
        rs("HireDate") = DateSerial(2018 + (i Mod 6), 1 + i, 1)
        rs("BasicSalary") = arr(i)(5)
        rs("AnnualLeaveDays") = arr(i)(6)
        rs("Phone") = arr(i)(7)
        rs("Active") = True
        rs("Notes") = arr(i)(8)
        rs.Update
    Next
    rs.Close
    LogErr "Employees data"
End Sub

'==========================================================================
' 3) Calendar, holidays and the sample month (previous month)
'==========================================================================
Sub AddAtt(rs, empId, d, tin, tout)
    On Error Resume Next
    rs.AddNew
    rs("EmpID") = empId
    rs("WorkDate") = d
    rs("TimeIn") = tin
    rs("TimeOut") = tout
    rs.Update
    LogErr "Attendance " & empId & " " & d
End Sub

Sub AddLeave(empId, typeId, d1, d2, note)
    On Error Resume Next
    Dim rs
    Set rs = db.OpenRecordset("tblLeaves", dbOpenDynaset)
    rs.AddNew
    rs("EmpID") = empId
    rs("LeaveTypeID") = typeId
    rs("FromDate") = d1
    rs("ToDate") = d2
    rs("Notes") = note
    rs.Update
    rs.Close
    LogErr "Leave " & empId
End Sub

Sub FillCalendarHolidaysAndSampleData()
    On Error Resume Next
    Dim rs, d, y0, y, i, k, n, hol, fixedH, wd(), mStart, mEnd, hd, tin, tout, fri1, sat2

    y0 = Year(Date)
    Set rs = db.OpenRecordset("tblCalendar", dbOpenDynaset)
    d = DateSerial(y0 - 1, 1, 1)
    Do While d <= DateSerial(y0 + 2, 12, 31)
        rs.AddNew
        rs("CalDate") = d
        rs.Update
        d = d + 1
    Loop
    rs.Close
    LogErr "Calendar data"

    ' fixed-date official holidays (edit freely from the holidays form)
    Set hol = CreateObject("Scripting.Dictionary")
    fixedH = Array(Array(1, 7, "عيد الميلاد المجيد"), Array(1, 25, "عيد الشرطة وثورة 25 يناير"), _
                   Array(4, 25, "عيد تحرير سيناء"), Array(5, 1, "عيد العمال"), Array(6, 30, "ذكرى ثورة 30 يونيو"), _
                   Array(7, 23, "عيد ثورة 23 يوليو"), Array(10, 6, "عيد القوات المسلحة"))
    For y = y0 - 1 To y0 + 2
        For i = 0 To UBound(fixedH)
            hol(CLng(DateSerial(y, fixedH(i)(0), fixedH(i)(1)))) = fixedH(i)(2)
        Next
    Next

    ' sample month = previous month; collect weekdays that are not holidays
    mStart = DateSerial(Year(Date), Month(Date) - 1, 1)
    mEnd = DateSerial(Year(mStart), Month(mStart) + 1, 0)
    n = 0
    ReDim wd(40)
    d = mStart
    Do While d <= mEnd
        If Weekday(d) <> 6 And Weekday(d) <> 7 And Not hol.Exists(CLng(d)) Then
            wd(n) = d
            n = n + 1
        End If
        d = d + 1
    Loop
    ' one demo holiday inside the sample month (used by case 8)
    hd = wd(13)
    hol(CLng(hd)) = "عطلة رسمية - مثال تجريبي (يمكن حذفها)"
    For i = 13 To n - 2
        wd(i) = wd(i + 1)
    Next
    n = n - 1

    Set rs = db.OpenRecordset("tblHolidays", dbOpenDynaset)
    For Each k In hol.Keys
        rs.AddNew
        rs("HolidayDate") = CDate(k)
        rs("HolidayName") = hol(k)
        rs.Update
    Next
    rs.Close
    LogErr "Holidays data"

    ' first Friday and second Saturday of the sample month (case 8 works them)
    fri1 = mStart
    Do While Weekday(fri1) <> 6
        fri1 = fri1 + 1
    Loop
    sat2 = fri1 + 8

    Set rs = db.OpenRecordset("tblAttendance", dbOpenDynaset)
    For i = 0 To n - 1
        d = wd(i)
        k = Day(d)

        ' 1001 punctual; business mission on wd(4)
        If i <> 4 Then AddAtt rs, 1001, d, T(8, 50 + (k * 3) Mod 9), T(17, k Mod 10)

        ' 1002 frequent lateness + early leave; unpaid leave on wd(9)
        If i <> 9 Then
            If k Mod 3 = 0 Then
                tin = T(9, 25 + (k Mod 5))
            Else
                tin = T(9, k Mod 10)
            End If
            If k Mod 4 = 0 Then
                tout = T(16, 40)
            Else
                tout = T(17, 5)
            End If
            AddAtt rs, 1002, d, tin, tout
        End If

        ' 1003 daytime overtime every other day
        If i Mod 2 = 1 Then
            tout = T(19, (k Mod 4) * 5)
        Else
            tout = T(17, 10)
        End If
        AddAtt rs, 1003, d, T(8, 55), tout

        ' 1004 night overtime on Mondays and Wednesdays
        If Weekday(d) = 2 Or Weekday(d) = 4 Then
            tout = T(22, 30)
        Else
            tout = T(17, 15)
        End If
        AddAtt rs, 1004, d, T(9, 0), tout

        ' 1005 annual leave wd(5)..wd(9)
        If i < 5 Or i > 9 Then AddAtt rs, 1005, d, T(8, 58), T(17, 2)

        ' 1006 absent without permission on wd(2), wd(10), wd(16)
        If i <> 2 And i <> 10 And i <> 16 Then AddAtt rs, 1006, d, T(9, 5), T(17, 0)

        ' 1007 casual leave wd(3), sick leave wd(11)..wd(14)
        If i <> 3 And (i < 11 Or i > 14) Then AddAtt rs, 1007, d, T(8, 57), T(17, 4)

        ' 1008 regular days
        AddAtt rs, 1008, d, T(8, 59), T(17, 3)

        ' 1009 early leave + one missing check-out
        If i = 8 Then
            tout = Null
        ElseIf k Mod 3 = 0 Then
            tout = T(16, 0)
        Else
            tout = T(17, 0)
        End If
        AddAtt rs, 1009, d, T(9, 10), tout

        ' 1010 evening shift 14:00-22:00, overtime sometimes past midnight
        If i Mod 5 = 4 Then
            tout = T(0, 40)
        ElseIf i Mod 2 = 0 Then
            tout = T(23, 45)
        Else
            tout = T(22, 5)
        End If
        AddAtt rs, 1010, d, T(13, 55), tout
    Next
    ' case 8: works the demo holiday and two weekend days
    AddAtt rs, 1008, hd, T(9, 0), T(17, 0)
    AddAtt rs, 1008, fri1, T(10, 0), T(16, 0)
    AddAtt rs, 1008, sat2, T(10, 0), T(18, 30)
    rs.Close

    AddLeave 1001, 5, wd(4), wd(4), "مأمورية لدى أحد العملاء"
    AddLeave 1002, 4, wd(9), wd(9), "ظروف خاصة"
    AddLeave 1005, 1, wd(5), wd(9), "إجازة اعتيادية مجدولة"
    AddLeave 1007, 2, wd(3), wd(3), "إجازة عارضة"
    AddLeave 1007, 3, wd(11), wd(14), "شهادة طبية"
End Sub

'==========================================================================
' 4) Queries (calculation engine + report sources)
'==========================================================================
Sub CreateQueries()
    Dim s, otMin

    MakeQ "qryEmp", "SELECT E.EmpID, E.EmpName, E.JobTitle, E.DeptID, D.DeptName, E.ShiftID, S.ShiftName, E.BasicSalary, " & _
        "E.HireDate, E.AnnualLeaveDays, E.Active FROM (tblEmployees AS E LEFT JOIN tblDepartments AS D ON E.DeptID = D.DeptID) " & _
        "LEFT JOIN tblShifts AS S ON E.ShiftID = S.ShiftID"

    MakeQ "qryEmpF", "PARAMETERS " & P_EMP & " Long, " & P_DEPT & " Long; SELECT Q.* FROM qryEmp AS Q " & _
        "WHERE (" & P_EMP & " Is Null Or Q.EmpID=" & P_EMP & ") AND (" & P_DEPT & " Is Null Or Q.DeptID=" & P_DEPT & ")"

    ' ---- attendance calculation chain ----
    MakeQ "qryAtt_Base", "SELECT A.AttID, A.EmpID, E.EmpName, D.DeptName, E.BasicSalary, A.WorkDate, A.TimeIn, A.TimeOut, A.Notes, " & _
        "S.ShiftName, S.StartTime, " & _
        "IIf(S.EndTime<=S.StartTime,S.EndTime+1,S.EndTime) AS ShiftEndAdj, " & _
        "IIf(S.LunchAfter<S.StartTime,S.LunchAfter+1,S.LunchAfter) AS LunchAdj, " & _
        "IIf(S.DinnerAfter<S.StartTime,S.DinnerAfter+1,S.DinnerAfter) AS DinnerAdj, " & _
        "IIf(A.TimeIn Is Null Or A.TimeOut Is Null,0,1) AS HasOut, " & _
        "IIf(A.TimeIn Is Null Or A.TimeOut Is Null,S.StartTime,IIf(A.TimeOut<A.TimeIn,A.TimeOut+1,A.TimeOut)) AS OutAdj, " & _
        "IIf(H.HolidayDate Is Null,0,1) AS IsHoliday " & _
        "FROM (((tblAttendance AS A INNER JOIN tblEmployees AS E ON A.EmpID = E.EmpID) " & _
        "INNER JOIN tblShifts AS S ON E.ShiftID = S.ShiftID) " & _
        "LEFT JOIN tblDepartments AS D ON E.DeptID = D.DeptID) " & _
        "LEFT JOIN tblHolidays AS H ON A.WorkDate = H.HolidayDate"

    MakeQ "qryAtt_Calc1", "SELECT B.*, " & _
        "IIf(Weekday(B.WorkDate)=G.WeekendDay1 Or Weekday(B.WorkDate)=G.WeekendDay2 Or B.IsHoliday=1,1,0) AS IsOff, " & _
        "IIf(B.HasOut=0,0,Round((B.OutAdj-B.TimeIn)*1440,0)) AS WorkedMin, " & _
        "Round((B.ShiftEndAdj-B.StartTime)*1440,0) AS ShiftMin, " & _
        "B.ShiftEndAdj+G.OTStartAfterMin/1440 AS OTStartT, " & _
        "IIf(G.NightStart<B.StartTime,G.NightStart+1,G.NightStart) AS NightAdj, " & _
        "G.GraceMinutes, G.OTMinMinutes, G.OTRateDay, G.OTRateNight, G.OTRateOff, G.LunchAmount, G.DinnerAmount, " & _
        "G.WorkDaysPerMonth, G.LateDeductOn " & _
        "FROM qryAtt_Base AS B, tblSettings AS G"

    MakeQ "qryAtt_Calc2", "SELECT C.*, " & _
        "IIf(C.IsOff=1 Or C.TimeIn Is Null,0,IIf(Round((C.TimeIn-C.StartTime)*1440,0)>C.GraceMinutes,Round((C.TimeIn-C.StartTime)*1440,0),0)) AS LateMin, " & _
        "IIf(C.IsOff=1 Or C.HasOut=0,0,IIf(C.OutAdj<C.ShiftEndAdj,Round((C.ShiftEndAdj-C.OutAdj)*1440,0),0)) AS EarlyMin, " & _
        "IIf(C.IsOff=1 Or C.HasOut=0,0,IIf(C.OutAdj>C.OTStartT,Round((C.OutAdj-C.OTStartT)*1440,0),0)) AS OTRawMin, " & _
        "IIf(C.IsOff=1 Or C.HasOut=0,0,IIf(C.OutAdj>IIf(C.NightAdj>C.OTStartT,C.NightAdj,C.OTStartT)," & _
        "Round((C.OutAdj-IIf(C.NightAdj>C.OTStartT,C.NightAdj,C.OTStartT))*1440,0),0)) AS NightRawMin, " & _
        "IIf(C.IsOff=1,C.WorkedMin,0) AS OffOTMin " & _
        "FROM qryAtt_Calc1 AS C"

    MakeQ "qryAtt_Calc3", "SELECT X.*, " & _
        "IIf(X.OTRawMin>0 And X.OTRawMin>=X.OTMinMinutes,X.OTRawMin-X.NightRawMin,0) AS DayOTMin, " & _
        "IIf(X.OTRawMin>0 And X.OTRawMin>=X.OTMinMinutes,X.NightRawMin,0) AS NightOTMin " & _
        "FROM qryAtt_Calc2 AS X"

    otMin = "(Y.DayOTMin+Y.NightOTMin+Y.OffOTMin)"
    MakeQ "qryAttCalc", "SELECT Y.*, 1 AS AllGrp, " & DayNameExpr("Y.WorkDate") & " AS DayName, " & _
        "Round(Y.WorkedMin/60,2) AS WorkedHrs, " & _
        "Round(Y.DayOTMin/60,2) AS DayOTHrs, Round(Y.NightOTMin/60,2) AS NightOTHrs, Round(Y.OffOTMin/60,2) AS OffOTHrs, " & _
        otMin & " AS OTMin, Round(" & otMin & "/60,2) AS OTHrs, " & _
        "Round(Y.BasicSalary/Y.WorkDaysPerMonth/(Y.ShiftMin/60),2) AS HourRate, " & _
        "Round(Y.BasicSalary/Y.WorkDaysPerMonth/(Y.ShiftMin/60)*(Y.DayOTMin*Y.OTRateDay+Y.NightOTMin*Y.OTRateNight+Y.OffOTMin*Y.OTRateOff)/60,2) AS OTPay, " & _
        "IIf(Y.HasOut=1 And " & otMin & ">0 And Y.OutAdj>=Y.LunchAdj,Y.LunchAmount,0) AS LunchAllow, " & _
        "IIf(Y.HasOut=1 And " & otMin & ">0 And Y.OutAdj>=Y.DinnerAdj,Y.DinnerAmount,0) AS DinnerAllow, " & _
        "IIf(Y.HasOut=1 And " & otMin & ">0 And Y.OutAdj>=Y.LunchAdj,Y.LunchAmount,0)+IIf(Y.HasOut=1 And " & otMin & ">0 And Y.OutAdj>=Y.DinnerAdj,Y.DinnerAmount,0) AS MealTotal, " & _
        "IIf(Y.LateDeductOn=True,Round(Y.BasicSalary/Y.WorkDaysPerMonth/Y.ShiftMin*(Y.LateMin+Y.EarlyMin),2),0) AS LateDeduct, " & _
        "IIf(Y.HasOut=0,""بصمة ناقصة"",IIf(Y.IsOff=1,IIf(Y.IsHoliday=1,""عمل في عطلة رسمية"",""عمل في يوم الراحة"")," & _
        "IIf(Y.LateMin>0 And Y.EarlyMin>0,""تأخير وانصراف مبكر"",IIf(Y.LateMin>0,""متأخر"",IIf(Y.EarlyMin>0,""انصراف مبكر""," & _
        "IIf(" & otMin & ">0,""منتظم + إضافي"",""منتظم"")))))) AS DayStatus " & _
        "FROM qryAtt_Calc3 AS Y"

    ' ---- report 1, 2, 6, 7 sources ----
    MakeQ "qryR01_Attendance", PARAM_DATES & "SELECT A.* FROM qryAttCalc AS A INNER JOIN qryEmpF AS F ON A.EmpID = F.EmpID " & _
        "WHERE A.WorkDate Between " & P_FROM & " And " & P_TO
    MakeQ "qryR02_LateEarly", "SELECT R.* FROM qryR01_Attendance AS R WHERE R.LateMin>0 Or R.EarlyMin>0 Or R.HasOut=0"
    MakeQ "qryR06_Overtime", "SELECT R.* FROM qryR01_Attendance AS R WHERE R.OTMin>0"
    MakeQ "qryR07_Meals", "SELECT R.* FROM qryR01_Attendance AS R WHERE R.MealTotal>0"

    ' ---- absence (derived from calendar) ----
    MakeQ "qryR03_Absence", PARAM_DATES & "SELECT E.EmpID, E.EmpName, E.DeptName, C.CalDate, 1 AS AbsCnt, 1 AS AllGrp, " & _
        DayNameExpr("C.CalDate") & " AS DayName " & _
        "FROM tblCalendar AS C, qryEmpF AS E, tblSettings AS G " & _
        "WHERE C.CalDate Between " & P_FROM & " And " & P_TO & " AND C.CalDate<=Date() AND E.Active=True " & _
        "AND (E.HireDate Is Null Or C.CalDate>=E.HireDate) " & _
        "AND Weekday(C.CalDate)<>G.WeekendDay1 AND Weekday(C.CalDate)<>G.WeekendDay2 " & _
        "AND C.CalDate NOT IN (SELECT HolidayDate FROM tblHolidays) " & _
        "AND NOT EXISTS (SELECT AttID FROM tblAttendance AS A WHERE A.EmpID=E.EmpID AND A.WorkDate=C.CalDate) " & _
        "AND NOT EXISTS (SELECT LeaveID FROM tblLeaves AS L WHERE L.EmpID=E.EmpID AND C.CalDate Between L.FromDate And L.ToDate)"

    ' ---- leaves ----
    MakeQ "qryLeavesFull", "SELECT L.LeaveID, L.EmpID, E.EmpName, E.DeptName, L.LeaveTypeID, T.LeaveTypeName, T.IsPaid, " & _
        "L.FromDate, L.ToDate, L.Notes FROM (tblLeaves AS L INNER JOIN qryEmp AS E ON L.EmpID = E.EmpID) " & _
        "INNER JOIN tblLeaveTypes AS T ON L.LeaveTypeID = T.LeaveTypeID"
    MakeQ "qryLeaveDays", "SELECT C.CalDate, L.LeaveID, L.EmpID, L.EmpName, L.DeptName, L.LeaveTypeID, L.LeaveTypeName, L.IsPaid, " & _
        "L.FromDate, L.ToDate, L.Notes FROM tblCalendar AS C, qryLeavesFull AS L, tblSettings AS G " & _
        "WHERE C.CalDate>=L.FromDate AND C.CalDate<=L.ToDate " & _
        "AND Weekday(C.CalDate)<>G.WeekendDay1 AND Weekday(C.CalDate)<>G.WeekendDay2 " & _
        "AND C.CalDate NOT IN (SELECT HolidayDate FROM tblHolidays)"
    MakeQ "qryR04_Leaves", PARAM_DATES & "SELECT L.LeaveID, L.EmpID, L.EmpName, L.DeptName, L.LeaveTypeName, " & _
        "IIf(L.IsPaid,""مدفوعة"",""بدون أجر"") AS PaidText, L.FromDate, L.ToDate, L.Notes, Count(L.CalDate) AS DaysInRange, 1 AS AllGrp " & _
        "FROM qryLeaveDays AS L INNER JOIN qryEmpF AS F ON L.EmpID = F.EmpID " & _
        "WHERE L.CalDate Between " & P_FROM & " And " & P_TO & " " & _
        "GROUP BY L.LeaveID, L.EmpID, L.EmpName, L.DeptName, L.LeaveTypeName, L.IsPaid, L.FromDate, L.ToDate, L.Notes"
    MakeQ "qryLeaveUsedYear", "PARAMETERS " & P_TO & " DateTime; SELECT L.EmpID, L.LeaveTypeID, Count(L.CalDate) AS UsedDays " & _
        "FROM qryLeaveDays AS L WHERE Year(L.CalDate)=Year(" & P_TO & ") GROUP BY L.EmpID, L.LeaveTypeID"
    MakeQ "qryBalBase", "SELECT E.EmpID, E.EmpName, E.DeptName, T.LeaveTypeID, T.LeaveTypeName, " & _
        "IIf(T.UseEmpBalance=True,E.AnnualLeaveDays,T.DefaultBalance) AS Entitled " & _
        "FROM qryEmpF AS E, tblLeaveTypes AS T WHERE (T.UseEmpBalance=True Or T.DefaultBalance>0) AND E.Active=True"
    MakeQ "qryR05_LeaveBalance", "SELECT B.EmpID, B.EmpName, B.DeptName, B.LeaveTypeID, B.LeaveTypeName, B.Entitled, " & _
        NZ0("U.UsedDays") & " AS UsedD, B.Entitled-" & NZ0("U.UsedDays") & " AS RemainD, 1 AS AllGrp " & _
        "FROM qryBalBase AS B LEFT JOIN qryLeaveUsedYear AS U ON (B.EmpID = U.EmpID) AND (B.LeaveTypeID = U.LeaveTypeID)"

    ' ---- monthly summary & payroll ----
    MakeQ "qryS_Att", "SELECT R.EmpID, Sum(IIf(R.IsOff=0,1,0)) AS sPresent, Sum(R.IsOff) AS sOffWork, Sum(R.WorkedMin) AS sWorkedMin, " & _
        "Sum(R.LateMin) AS sLateMin, Sum(R.EarlyMin) AS sEarlyMin, Sum(IIf(R.HasOut=0,1,0)) AS sMissing, " & _
        "Sum(R.DayOTMin) AS sDayOT, Sum(R.NightOTMin) AS sNightOT, Sum(R.OffOTMin) AS sOffOT, Sum(R.OTMin) AS sOTMin, " & _
        "Sum(R.OTPay) AS sOTPay, Sum(R.LunchAllow) AS sLunch, Sum(R.DinnerAllow) AS sDinner, Sum(R.LateDeduct) AS sLateDed " & _
        "FROM qryR01_Attendance AS R GROUP BY R.EmpID"
    MakeQ "qryS_Abs", "SELECT A.EmpID, Count(A.CalDate) AS sAbsDays FROM qryR03_Absence AS A GROUP BY A.EmpID"
    MakeQ "qryS_Leave", PARAM_DATES & "SELECT L.EmpID, Sum(IIf(L.IsPaid,1,0)) AS sPaidLeave, Sum(IIf(L.IsPaid,0,1)) AS sUnpaidLeave " & _
        "FROM qryLeaveDays AS L WHERE L.CalDate Between " & P_FROM & " And " & P_TO & " GROUP BY L.EmpID"

    MakeQ "qryR08_Summary", "SELECT E.EmpID, E.EmpName, E.DeptName, E.JobTitle, E.BasicSalary, 1 AS AllGrp, " & _
        NZ0("A.sPresent") & " AS PresentDays, " & NZ0("A.sOffWork") & " AS OffWorkDays, " & NZ0("B.sAbsDays") & " AS AbsDays, " & _
        NZ0("L.sPaidLeave") & " AS PaidLeaveDays, " & NZ0("L.sUnpaidLeave") & " AS UnpaidLeaveDays, " & _
        NZ0("A.sLateMin") & " AS LateMin, " & NZ0("A.sEarlyMin") & " AS EarlyMin, " & NZ0("A.sMissing") & " AS MissingPunch, " & _
        "Round(" & NZ0("A.sWorkedMin") & "/60,2) AS WorkedHrs, Round(" & NZ0("A.sDayOT") & "/60,2) AS DayOTHrs, " & _
        "Round(" & NZ0("A.sNightOT") & "/60,2) AS NightOTHrs, Round(" & NZ0("A.sOffOT") & "/60,2) AS OffOTHrs, " & _
        "Round(" & NZ0("A.sOTMin") & "/60,2) AS OTHrs, " & NZ0("A.sOTPay") & " AS OTPay, " & _
        NZ0("A.sLunch") & " AS LunchAllow, " & NZ0("A.sDinner") & " AS DinnerAllow, " & NZ0("A.sLateDed") & " AS LateDeduct " & _
        "FROM ((qryEmpF AS E LEFT JOIN qryS_Att AS A ON E.EmpID = A.EmpID) " & _
        "LEFT JOIN qryS_Abs AS B ON E.EmpID = B.EmpID) LEFT JOIN qryS_Leave AS L ON E.EmpID = L.EmpID " & _
        "WHERE E.Active=True"

    MakeQ "qryPayroll", "SELECT S.*, Round(S.BasicSalary/G.WorkDaysPerMonth,2) AS DailyRate, " & _
        "Round(S.AbsDays*G.AbsenceDeductDays*S.BasicSalary/G.WorkDaysPerMonth,2) AS AbsDeduct, " & _
        "Round(S.UnpaidLeaveDays*S.BasicSalary/G.WorkDaysPerMonth,2) AS UnpaidDeduct " & _
        "FROM qryR08_Summary AS S, tblSettings AS G"

    MakeQ "qryR09_Payroll", "SELECT P.*, P.LunchAllow+P.DinnerAllow AS MealTotal, " & _
        "P.BasicSalary+P.OTPay+P.LunchAllow+P.DinnerAllow AS Gross, " & _
        "P.AbsDeduct+P.UnpaidDeduct+P.LateDeduct AS TotalDeduct, " & _
        "(P.BasicSalary+P.OTPay+P.LunchAllow+P.DinnerAllow)-(P.AbsDeduct+P.UnpaidDeduct+P.LateDeduct) AS NetPay " & _
        "FROM qryPayroll AS P"

    MakeQ "qryR10_Dept", "SELECT P.DeptName, 1 AS AllGrp, Count(P.EmpID) AS dEmpCount, Sum(P.PresentDays) AS dPresent, " & _
        "Sum(P.AbsDays) AS dAbs, Sum(P.PaidLeaveDays+P.UnpaidLeaveDays) AS dLeave, Round(Sum(P.LateMin)/60,2) AS dLateHrs, " & _
        "Sum(P.OTHrs) AS dOTHrs, Sum(P.OTPay) AS dOTPay, Sum(P.MealTotal) AS dMeals, Sum(P.BasicSalary) AS dBasic, " & _
        "Sum(P.TotalDeduct) AS dDeduct, Sum(P.NetPay) AS dNet FROM qryR09_Payroll AS P GROUP BY P.DeptName"
End Sub

'==========================================================================
' 5) Forms
'==========================================================================
Function NewCtl(frmName, cType, sect, src, l, t, w, h)
    On Error Resume Next
    Set NewCtl = Nothing
    Set NewCtl = app.CreateControl(frmName, cType, sect, "", src, CLng(l), CLng(t), CLng(w), CLng(h))
    LogErr "CreateControl " & frmName & " " & src
End Function

' flds: Array(Array(field, caption, widthCm, kind[T|C|K], rowSource, format, default))
Sub BuildDsForm(fName, fCaption, recSrc, flds, readOnly)
    On Error Resume Next
    Dim frm, tmp, i, f, ctl, cType
    Set frm = app.CreateForm()
    tmp = frm.Name
    LogErr "CreateForm " & fName
    frm.RecordSource = recSrc
    frm.Caption = fCaption
    frm.DefaultView = 2
    frm.Orientation = 1
    If readOnly Then
        frm.AllowEdits = False
        frm.AllowAdditions = False
        frm.AllowDeletions = False
    End If
    LogErr "Form props " & fName
    For i = 0 To UBound(flds)
        f = flds(i)
        If f(3) = "C" Then
            cType = acComboBox
        ElseIf f(3) = "K" Then
            cType = acCheckBox
        Else
            cType = acTextBox
        End If
        Set ctl = NewCtl(tmp, cType, acDetail, f(0), 3 * CM, 0.2 * CM + i * 0.7 * CM, f(2) * CM, 0.6 * CM)
        LogErr "Control " & fName & "." & f(0)
        ctl.Name = SafeName(f(1))
        ctl.FontName = FONT_NAME
        If f(3) = "C" Then
            ctl.RowSource = f(4)
            ctl.ColumnCount = 2
            ctl.ColumnWidths = "0;" & CLng(f(2) * CM)
            ctl.BoundColumn = 1
            ctl.ListWidth = CLng(f(2) * CM)
            ctl.LimitToList = True
        End If
        If f(5) <> "" Then ctl.Format = f(5)
        If f(6) <> "" Then ctl.DefaultValue = f(6)
        LogErr "Control props " & fName & "." & f(0)
        ctl.DatasheetCaption = f(1)
        ctl.ColumnWidth = CLng(f(2) * CM)
        Err.Clear
    Next
    app.DoCmd.Close acForm, tmp, acSaveYes
    app.DoCmd.Rename fName, acForm, tmp
    LogErr "Save form " & fName
End Sub

Sub CreateForms()
    Dim rsEmp, rsDept, rsShift, rsLT
    rsEmp = "SELECT EmpID, EmpName FROM tblEmployees ORDER BY EmpID"
    rsDept = "SELECT DeptID, DeptName FROM tblDepartments ORDER BY DeptName"
    rsShift = "SELECT ShiftID, ShiftName FROM tblShifts ORDER BY ShiftID"
    rsLT = "SELECT LeaveTypeID, LeaveTypeName FROM tblLeaveTypes ORDER BY LeaveTypeID"

    BuildDsForm "frmAttendance", "تسجيل الحضور والانصراف", "SELECT * FROM tblAttendance ORDER BY WorkDate DESC, EmpID", Array( _
        Array("EmpID", "الموظف", 5, "C", rsEmp, "", ""), _
        Array("WorkDate", "التاريخ", 2.8, "T", "", "dd/mm/yyyy", "=Date()"), _
        Array("TimeIn", "وقت الحضور", 2.4, "T", "", "hh:nn", ""), _
        Array("TimeOut", "وقت الانصراف", 2.4, "T", "", "hh:nn", ""), _
        Array("Notes", "ملاحظات", 6, "T", "", "", "")), False

    BuildDsForm "frmLeaves", "تسجيل الإجازات والمأموريات", "SELECT * FROM tblLeaves ORDER BY FromDate DESC", Array( _
        Array("EmpID", "الموظف", 5, "C", rsEmp, "", ""), _
        Array("LeaveTypeID", "نوع الإجازة", 4.5, "C", rsLT, "", ""), _
        Array("FromDate", "من تاريخ", 2.8, "T", "", "dd/mm/yyyy", "=Date()"), _
        Array("ToDate", "إلى تاريخ", 2.8, "T", "", "dd/mm/yyyy", "=Date()"), _
        Array("Notes", "ملاحظات", 6, "T", "", "", "")), False

    BuildDsForm "frmEmployees", "بيانات الموظفين", "SELECT * FROM tblEmployees ORDER BY EmpID", Array( _
        Array("EmpID", "كود الموظف", 2.2, "T", "", "", ""), _
        Array("EmpName", "اسم الموظف", 5, "T", "", "", ""), _
        Array("DeptID", "الإدارة", 4, "C", rsDept, "", ""), _
        Array("JobTitle", "الوظيفة", 3.5, "T", "", "", ""), _
        Array("ShiftID", "الوردية", 3.5, "C", rsShift, "", ""), _
        Array("HireDate", "تاريخ التعيين", 2.6, "T", "", "dd/mm/yyyy", ""), _
        Array("BasicSalary", "الأجر الأساسي", 2.8, "T", "", "#,##0.00", ""), _
        Array("AnnualLeaveDays", "رصيد الاعتيادي", 2.4, "T", "", "", ""), _
        Array("Phone", "الهاتف", 3, "T", "", "", ""), _
        Array("Active", "على رأس العمل", 2.2, "K", "", "", ""), _
        Array("Notes", "ملاحظات", 7, "T", "", "", "")), False

    BuildDsForm "frmShifts", "الورديات وساعات العمل الأساسية", "SELECT * FROM tblShifts ORDER BY ShiftID", Array( _
        Array("ShiftName", "اسم الوردية", 4.5, "T", "", "", ""), _
        Array("StartTime", "بداية الدوام", 2.6, "T", "", "hh:nn", ""), _
        Array("EndTime", "نهاية الدوام", 2.6, "T", "", "hh:nn", ""), _
        Array("LunchAfter", "بدل الغداء بعد الساعة", 3.4, "T", "", "hh:nn", ""), _
        Array("DinnerAfter", "بدل العشاء بعد الساعة", 3.4, "T", "", "hh:nn", "")), False

    BuildDsForm "frmDepartments", "الإدارات", "SELECT * FROM tblDepartments ORDER BY DeptID", Array( _
        Array("DeptName", "اسم الإدارة", 7, "T", "", "", "")), False

    BuildDsForm "frmLeaveTypes", "أنواع الإجازات", "SELECT * FROM tblLeaveTypes ORDER BY LeaveTypeID", Array( _
        Array("LeaveTypeName", "نوع الإجازة", 5, "T", "", "", ""), _
        Array("IsPaid", "مدفوعة الأجر", 2.4, "K", "", "", ""), _
        Array("DefaultBalance", "الرصيد السنوي", 2.6, "T", "", "", ""), _
        Array("UseEmpBalance", "يُستخدم رصيد الموظف", 3.2, "K", "", "", "")), False

    BuildDsForm "frmHolidays", "العطلات الرسمية", "SELECT * FROM tblHolidays ORDER BY HolidayDate", Array( _
        Array("HolidayDate", "تاريخ العطلة", 3, "T", "", "dd/mm/yyyy", ""), _
        Array("HolidayName", "اسم العطلة", 7, "T", "", "", "")), False

    BuildDsForm "frmAttCalc", "عرض الحسابات اليومية (للقراءة فقط)", "SELECT * FROM qryAttCalc ORDER BY WorkDate DESC, EmpID", Array( _
        Array("EmpName", "الموظف", 4.2, "T", "", "", ""), _
        Array("WorkDate", "التاريخ", 2.4, "T", "", "dd/mm/yyyy", ""), _
        Array("DayName", "اليوم", 2, "T", "", "", ""), _
        Array("TimeIn", "حضور", 1.7, "T", "", "hh:nn", ""), _
        Array("TimeOut", "انصراف", 1.7, "T", "", "hh:nn", ""), _
        Array("WorkedHrs", "ساعات العمل", 2, "T", "", "0.00", ""), _
        Array("LateMin", "تأخير د", 1.6, "T", "", "", ""), _
        Array("EarlyMin", "انصراف مبكر د", 2.2, "T", "", "", ""), _
        Array("DayOTHrs", "إضافي نهاري س", 2.2, "T", "", "0.00", ""), _
        Array("NightOTHrs", "إضافي ليلي س", 2.2, "T", "", "0.00", ""), _
        Array("OffOTHrs", "إضافي راحات س", 2.2, "T", "", "0.00", ""), _
        Array("OTPay", "قيمة الإضافي", 2.2, "T", "", "#,##0.00", ""), _
        Array("LunchAllow", "بدل غداء", 1.8, "T", "", "#,##0.00", ""), _
        Array("DinnerAllow", "بدل عشاء", 1.8, "T", "", "#,##0.00", ""), _
        Array("DayStatus", "الحالة", 3.5, "T", "", "", "")), True

    BuildSettingsForm
    BuildMainForm
End Sub

Sub AddLabel(frmName, sect, cap, l, t, w, h, fs, bold)
    On Error Resume Next
    Dim c
    Set c = NewCtl(frmName, acLabel, sect, "", l, t, w, h)
    c.Caption = cap
    c.FontName = FONT_NAME
    c.FontSize = fs
    c.FontBold = bold
    LogErr "Label " & cap
End Sub

Sub BuildSettingsForm()
    On Error Resume Next
    Dim frm, tmp, i, f, ctl, col, r, x, y, flds, wdList
    wdList = "1;الأحد;2;الاثنين;3;الثلاثاء;4;الأربعاء;5;الخميس;6;الجمعة;7;السبت"
    flds = Array( _
        Array("CompanyName", "اسم الشركة", "T", ""), _
        Array("CurrencyName", "العملة", "T", ""), _
        Array("GraceMinutes", "سماحية التأخير (دقيقة)", "T", ""), _
        Array("OTStartAfterMin", "بدء الإضافي بعد نهاية الوردية (دقيقة)", "T", ""), _
        Array("OTMinMinutes", "الحد الأدنى لاحتساب الإضافي (دقيقة)", "T", ""), _
        Array("NightStart", "بداية احتساب الإضافي الليلي", "T", "hh:nn"), _
        Array("OTRateDay", "معامل الإضافي النهاري (مثال 1.35)", "T", "0.00"), _
        Array("OTRateNight", "معامل الإضافي الليلي (مثال 1.70)", "T", "0.00"), _
        Array("OTRateOff", "معامل العمل في الراحة/العطلة (مثال 2)", "T", "0.00"), _
        Array("LunchAmount", "قيمة بدل الغداء", "T", "#,##0.00"), _
        Array("DinnerAmount", "قيمة بدل العشاء", "T", "#,##0.00"), _
        Array("WorkDaysPerMonth", "عدد أيام الشهر لحساب الأجر اليومي", "T", ""), _
        Array("WeekendDay1", "يوم الراحة الأسبوعية الأول", "W", ""), _
        Array("WeekendDay2", "يوم الراحة الأسبوعية الثاني", "W", ""), _
        Array("AbsenceDeductDays", "أيام الخصم عن كل يوم غياب", "T", "0.00"), _
        Array("LateDeductOn", "تفعيل خصم التأخير والانصراف المبكر", "K", ""))

    Set frm = app.CreateForm()
    tmp = frm.Name
    frm.RecordSource = "tblSettings"
    frm.Caption = "الإعدادات العامة وقواعد الحساب"
    frm.DefaultView = 0
    frm.Orientation = 1
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.AutoCenter = True
    frm.Width = 19 * CM
    frm.Section(acDetail).Height = 10.5 * CM
    LogErr "Settings form props"

    AddLabel tmp, acDetail, "الإعدادات العامة وقواعد حساب الإضافي والبدلات", 0.3 * CM, 0.2 * CM, 18.4 * CM, 0.9 * CM, 14, True
    For i = 0 To UBound(flds)
        f = flds(i)
        col = i \ 8
        r = i Mod 8
        x = 0.3 * CM + col * 9.4 * CM
        y = 1.4 * CM + r * 0.85 * CM
        AddLabel tmp, acDetail, f(1), x, y, 5.6 * CM, 0.7 * CM, 10, False
        If f(2) = "K" Then
            Set ctl = NewCtl(tmp, acCheckBox, acDetail, f(0), x + 6 * CM, y + 0.1 * CM, 0.6 * CM, 0.5 * CM)
        ElseIf f(2) = "W" Then
            Set ctl = NewCtl(tmp, acComboBox, acDetail, f(0), x + 6 * CM, y, 2.9 * CM, 0.65 * CM)
            ctl.RowSourceType = "Value List"
            ctl.RowSource = wdList
            ctl.ColumnCount = 2
            ctl.ColumnWidths = "0;" & CLng(2.9 * CM)
            ctl.BoundColumn = 1
            ctl.LimitToList = True
        Else
            Set ctl = NewCtl(tmp, acTextBox, acDetail, f(0), x + 6 * CM, y, 2.9 * CM, 0.65 * CM)
            If f(3) <> "" Then ctl.Format = f(3)
        End If
        ctl.Name = "c_" & f(0)
        ctl.FontName = FONT_NAME
        LogErr "Settings ctl " & f(0)
    Next
    AddLabel tmp, acDetail, "ملاحظة: مواعيد بداية ونهاية الدوام ومواعيد استحقاق بدل الغداء والعشاء تُعدّل من شاشة (الورديات وساعات العمل) لكل وردية على حدة. " & _
        "قيمة ساعة الأجر = الأجر الأساسي ÷ أيام الشهر ÷ ساعات الوردية.", 0.3 * CM, 8.4 * CM, 18.4 * CM, 1.4 * CM, 9, False

    app.DoCmd.Close acForm, tmp, acSaveYes
    app.DoCmd.Rename "frmSettings", acForm, tmp
    LogErr "Save frmSettings"
End Sub

Sub AddButton(frmName, cName, cap, l, t, w, h)
    On Error Resume Next
    Dim c
    Set c = NewCtl(frmName, acCommandButton, acDetail, "", l, t, w, h)
    c.Name = cName
    c.Caption = cap
    c.FontName = FONT_NAME
    c.FontSize = 10
    c.FontBold = True
    c.OnClick = "[Event Procedure]"
    LogErr "Button " & cName
End Sub

Sub BuildMainForm()
    On Error Resume Next
    Dim frm, tmp, c, i, rpts, dataBtns, code, x, y

    Set frm = app.CreateForm()
    tmp = frm.Name
    frm.Caption = "نظام الحضور والانصراف والأجور الإضافية"
    frm.DefaultView = 0
    frm.Orientation = 1
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.AutoCenter = True
    frm.Width = 17 * CM
    frm.Section(acDetail).Height = 15.6 * CM
    frm.Section(acDetail).BackColor = RGB(244, 247, 251)
    LogErr "Main form props"

    Set c = NewCtl(tmp, acLabel, acDetail, "", 0.4 * CM, 0.3 * CM, 16.2 * CM, 1.1 * CM)
    c.Caption = "نظام الحضور والانصراف والأجور الإضافية"
    c.FontName = FONT_NAME
    c.FontSize = 18
    c.FontBold = True
    c.TextAlign = 2
    c.BackStyle = 1
    c.BackColor = RGB(31, 78, 121)
    c.ForeColor = RGB(255, 255, 255)
    Set c = NewCtl(tmp, acTextBox, acDetail, "", 0.4 * CM, 1.5 * CM, 16.2 * CM, 0.6 * CM)
    c.Name = "txtCompany"
    c.ControlSource = "=DLookUp(""CompanyName"",""tblSettings"",""ID=1"")"
    c.FontName = FONT_NAME
    c.FontSize = 11
    c.TextAlign = 2
    c.BorderStyle = 0
    c.BackStyle = 0
    c.Locked = True
    LogErr "Main header"

    ' criteria
    Set c = NewCtl(tmp, acRectangle, acDetail, "", 0.4 * CM, 2.3 * CM, 16.2 * CM, 2.2 * CM)
    AddLabel tmp, acDetail, "من تاريخ", 0.7 * CM, 2.55 * CM, 2 * CM, 0.6 * CM, 10, True
    Set c = NewCtl(tmp, acTextBox, acDetail, "", 2.8 * CM, 2.5 * CM, 3 * CM, 0.65 * CM)
    c.Name = "txtFrom"
    c.Format = "dd/mm/yyyy"
    c.DefaultValue = "=DateSerial(Year(Date()),Month(Date())-1,1)"
    c.FontName = FONT_NAME
    AddLabel tmp, acDetail, "إلى تاريخ", 6.3 * CM, 2.55 * CM, 2 * CM, 0.6 * CM, 10, True
    Set c = NewCtl(tmp, acTextBox, acDetail, "", 8.4 * CM, 2.5 * CM, 3 * CM, 0.65 * CM)
    c.Name = "txtTo"
    c.Format = "dd/mm/yyyy"
    c.DefaultValue = "=DateSerial(Year(Date()),Month(Date()),0)"
    c.FontName = FONT_NAME
    AddLabel tmp, acDetail, "الموظف", 0.7 * CM, 3.55 * CM, 2 * CM, 0.6 * CM, 10, True
    Set c = NewCtl(tmp, acComboBox, acDetail, "", 2.8 * CM, 3.5 * CM, 3.3 * CM, 0.65 * CM)
    c.Name = "cboEmp"
    c.RowSource = "SELECT EmpID, EmpName FROM tblEmployees ORDER BY EmpID"
    c.ColumnCount = 2
    c.ColumnWidths = "0;" & CLng(4.5 * CM)
    c.ListWidth = CLng(4.5 * CM)
    c.BoundColumn = 1
    c.LimitToList = True
    c.FontName = FONT_NAME
    AddLabel tmp, acDetail, "الإدارة", 6.3 * CM, 3.55 * CM, 2 * CM, 0.6 * CM, 10, True
    Set c = NewCtl(tmp, acComboBox, acDetail, "", 8.4 * CM, 3.5 * CM, 3 * CM, 0.65 * CM)
    c.Name = "cboDept"
    c.RowSource = "SELECT DeptID, DeptName FROM tblDepartments ORDER BY DeptName"
    c.ColumnCount = 2
    c.ColumnWidths = "0;" & CLng(4 * CM)
    c.ListWidth = CLng(4 * CM)
    c.BoundColumn = 1
    c.LimitToList = True
    c.FontName = FONT_NAME
    AddButton tmp, "cmdClear", "إلغاء فلتر الموظف/الإدارة", 11.8 * CM, 3.45 * CM, 4.5 * CM, 0.75 * CM
    AddLabel tmp, acDetail, "اترك الموظف والإدارة فارغين لعرض الجميع", 11.8 * CM, 2.55 * CM, 4.6 * CM, 0.6 * CM, 8, False
    LogErr "Main criteria"

    ' report buttons (2 columns x 5 rows)
    AddLabel tmp, acDetail, "التقارير", 0.4 * CM, 4.75 * CM, 16.2 * CM, 0.7 * CM, 13, True
    rpts = Array("01- الحضور والانصراف اليومي التفصيلي", "02- التأخير والانصراف المبكر والبصمات الناقصة", _
                 "03- الغياب بدون إذن", "04- الإجازات والمأموريات", "05- أرصدة الإجازات السنوية", _
                 "06- الساعات الإضافية وقيمتها", "07- بدلات الغداء والعشاء", "08- الملخص الشهري للحضور", _
                 "09- كشف الأجور الأساسية والإضافية", "10- ملخص الإدارات")
    For i = 0 To 9
        x = 0.4 * CM + (i Mod 2) * 8.2 * CM
        y = 5.5 * CM + (i \ 2) * 0.85 * CM
        AddButton tmp, "cmdR" & Right("0" & (i + 1), 2), rpts(i), x, y, 8 * CM, 0.75 * CM
    Next

    ' data buttons (3 columns x 4 rows)
    AddLabel tmp, acDetail, "إدخال البيانات والإعدادات", 0.4 * CM, 9.9 * CM, 16.2 * CM, 0.7 * CM, 13, True
    dataBtns = Array("تسجيل الحضور والانصراف", "تسجيل الإجازات والمأموريات", "بيانات الموظفين", _
                     "الإعدادات العامة", "الورديات وساعات العمل", "الإدارات", _
                     "أنواع الإجازات", "العطلات الرسمية", "عرض الحسابات اليومية", _
                     "إضافة سنة للتقويم", "تحديث البيانات", "خروج")
    For i = 0 To 11
        x = 0.4 * CM + (i Mod 3) * 5.45 * CM
        y = 10.65 * CM + (i \ 3) * 0.85 * CM
        AddButton tmp, "cmdD" & Right("0" & (i + 1), 2), dataBtns(i), x, y, 5.3 * CM, 0.75 * CM
    Next

    ' hidden Arabic messages used by the VBA code (keeps VBA code ASCII-only)
    AddHidden tmp, "lblMsgDates", "من فضلك أدخل تاريخ البداية وتاريخ النهاية", 0
    AddHidden tmp, "lblMsgOrder", "تاريخ البداية يجب أن يكون قبل تاريخ النهاية", 1
    AddHidden tmp, "lblMsgNoData", "لا توجد بيانات مطابقة للفترة/الفلتر المحدد", 2
    AddHidden tmp, "lblMsgCal", "تمت إضافة سنة جديدة للتقويم حتى تاريخ", 3

    code = "Private Sub OpenRpt(ByVal rptName As String)" & vbCrLf & _
        "    If IsNull(Me.txtFrom) Or IsNull(Me.txtTo) Then" & vbCrLf & _
        "        MsgBox Me.lblMsgDates.Caption, vbExclamation" & vbCrLf & _
        "        Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    If Me.txtFrom > Me.txtTo Then" & vbCrLf & _
        "        MsgBox Me.lblMsgOrder.Caption, vbExclamation" & vbCrLf & _
        "        Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    DoCmd.OpenReport rptName, acViewPreview" & vbCrLf & _
        "    Err.Clear" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub cmdClear_Click()" & vbCrLf & _
        "    Me.cboEmp = Null" & vbCrLf & _
        "    Me.cboDept = Null" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf

    rpts = Array("rpt01_DailyAttendance", "rpt02_LateEarly", "rpt03_Absence", "rpt04_Leaves", "rpt05_LeaveBalance", _
                 "rpt06_Overtime", "rpt07_Meals", "rpt08_MonthlySummary", "rpt09_Payroll", "rpt10_Departments")
    For i = 0 To 9
        code = code & "Private Sub cmdR" & Right("0" & (i + 1), 2) & "_Click()" & vbCrLf & _
            "    OpenRpt """ & rpts(i) & """" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    Next
    dataBtns = Array("frmAttendance", "frmLeaves", "frmEmployees", "frmSettings", "frmShifts", _
                     "frmDepartments", "frmLeaveTypes", "frmHolidays", "frmAttCalc")
    For i = 0 To 8
        code = code & "Private Sub cmdD" & Right("0" & (i + 1), 2) & "_Click()" & vbCrLf & _
            "    DoCmd.OpenForm """ & dataBtns(i) & """" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    Next
    code = code & "Private Sub cmdD10_Click()" & vbCrLf & _
        "    Dim d As Date, mx As Variant, i As Long, rs As Object" & vbCrLf & _
        "    mx = DMax(""CalDate"", ""tblCalendar"")" & vbCrLf & _
        "    If IsNull(mx) Then d = DateSerial(Year(Date) - 1, 1, 1) Else d = mx + 1" & vbCrLf & _
        "    Set rs = CurrentDb.OpenRecordset(""tblCalendar"", 2)" & vbCrLf & _
        "    For i = 1 To 365" & vbCrLf & _
        "        rs.AddNew" & vbCrLf & _
        "        rs!CalDate = d" & vbCrLf & _
        "        rs.Update" & vbCrLf & _
        "        d = d + 1" & vbCrLf & _
        "    Next i" & vbCrLf & _
        "    rs.Close" & vbCrLf & _
        "    MsgBox Me.lblMsgCal.Caption & "" "" & Format(d - 1, ""dd/mm/yyyy""), vbInformation" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub cmdD11_Click()" & vbCrLf & _
        "    Me.cboEmp.Requery" & vbCrLf & _
        "    Me.cboDept.Requery" & vbCrLf & _
        "    Me.txtCompany.Requery" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub cmdD12_Click()" & vbCrLf & _
        "    DoCmd.Quit" & vbCrLf & _
        "End Sub" & vbCrLf

    frm.HasModule = True
    frm.Module.AddFromString code
    LogErr "Main form code"

    app.DoCmd.Close acForm, tmp, acSaveYes
    app.DoCmd.Rename "frmMain", acForm, tmp
    LogErr "Save frmMain"
End Sub

Sub AddHidden(frmName, cName, cap, idx)
    On Error Resume Next
    Dim c
    Set c = NewCtl(frmName, acLabel, acDetail, "", 0.4 * CM + idx * 1.2 * CM, 14.9 * CM, 1 * CM, 0.4 * CM)
    c.Name = cName
    c.Caption = cap
    c.Visible = False
    LogErr "Hidden " & cName
End Sub

'==========================================================================
' 6) Reports
'==========================================================================
Function NewRCtl(rName, cType, sect, src, l, t, w, h)
    On Error Resume Next
    Set NewRCtl = Nothing
    Set NewRCtl = app.CreateReportControl(rName, cType, sect, "", src, CLng(l), CLng(t), CLng(w), CLng(h))
    LogErr "CreateReportControl " & rName & " " & src
End Function

Sub RLabel(rName, sect, cap, l, t, w, h, fs, bold, shade)
    On Error Resume Next
    Dim c
    Set c = NewRCtl(rName, acLabel, sect, "", l, t, w, h)
    c.Caption = cap
    c.FontName = FONT_NAME
    c.FontSize = fs
    c.FontBold = bold
    c.TextAlign = 2
    If shade Then
        c.BackStyle = 1
        c.BackColor = RGB(221, 235, 247)
        c.BorderStyle = 1
        c.BorderColor = RGB(150, 150, 150)
    End If
    LogErr "RLabel " & rName
End Sub

Function RText(rName, sect, src, l, t, w, h, fs, bold, fmt, border)
    On Error Resume Next
    Dim c
    Set c = NewRCtl(rName, acTextBox, sect, "", l, t, w, h)
    c.ControlSource = src
    c.FontName = FONT_NAME
    c.FontSize = fs
    c.FontBold = bold
    c.TextAlign = 2
    If fmt <> "" Then c.Format = fmt
    If border Then
        c.BorderStyle = 1
        c.BorderColor = RGB(190, 190, 190)
    Else
        c.BorderStyle = 0
    End If
    LogErr "RText " & rName & " " & src
    Set RText = c
End Function

' cols: Array(Array(field, caption, widthCm, format, sumFlag))
Sub BuildReport(rName, rCaption, recSrc, grpField, grpTitle, sortField, cols, landscape, periodExpr)
    On Error Resume Next
    Dim rpt, tmp, i, x, w, totalW, hasGrp, secCols, yCols, c, rowH, firstLabelDone
    Set rpt = app.CreateReport()
    tmp = rpt.Name
    LogErr "CreateReport " & rName
    rpt.RecordSource = recSrc
    rpt.Caption = rCaption
    rpt.Orientation = 1
    If landscape Then rpt.Printer.Orientation = 2
    Err.Clear

    totalW = 0
    For i = 0 To UBound(cols)
        totalW = totalW + CLng(cols(i)(2) * CM)
    Next
    rpt.Width = totalW + 30
    rowH = CLng(0.6 * CM)

    hasGrp = (grpField <> "")
    app.CreateGroupLevel tmp, "AllGrp", True, True
    If hasGrp Then app.CreateGroupLevel tmp, grpField, True, True
    If sortField <> "" Then app.CreateGroupLevel tmp, sortField, False, False
    LogErr "Groups " & rName

    ' page header not used; page footer = page numbers
    rpt.Section(acPageHeader).Height = 0
    Err.Clear
    rpt.Section(acPageFooter).Height = CLng(0.7 * CM)
    RText tmp, acPageFooter, "=""صفحة "" & [Page] & "" من "" & [Pages] & ""   -   طُبع في "" & Format(Now(),""dd/mm/yyyy hh:nn"")", _
          0, CLng(0.1 * CM), totalW, CLng(0.5 * CM), 8, False, "", False
    Err.Clear

    ' report title block (repeats on every page)
    RText tmp, 5, "=DLookUp(""CompanyName"",""tblSettings"",""ID=1"")", 0, 0, totalW, CLng(0.6 * CM), 11, True, "", False
    RLabel tmp, 5, rCaption, 0, CLng(0.65 * CM), totalW, CLng(0.9 * CM), 16, True, False
    RText tmp, 5, periodExpr, 0, CLng(1.6 * CM), totalW, CLng(0.6 * CM), 10, False, "", False
    rpt.Section(5).RepeatSection = True
    LogErr "Title " & rName

    If hasGrp Then
        Set c = RText(tmp, 7, grpTitle, 0, CLng(0.15 * CM), totalW, CLng(0.65 * CM), 11, True, "", False)
        c.BackStyle = 1
        c.BackColor = RGB(242, 242, 242)
        c.TextAlign = 3
        rpt.Section(7).RepeatSection = True
        secCols = 7
        yCols = CLng(0.9 * CM)
        rpt.Section(5).Height = CLng(2.4 * CM)
    Else
        secCols = 5
        yCols = CLng(2.4 * CM)
    End If
    LogErr "Group header " & rName

    ' column headings, detail, group totals, grand totals
    x = 0
    firstLabelDone = False
    For i = 0 To UBound(cols)
        w = CLng(cols(i)(2) * CM)
        RLabel tmp, secCols, cols(i)(1), x, yCols, w, CLng(0.75 * CM), 9, True, True
        RText tmp, acDetail, cols(i)(0), x, 0, w, rowH, 9, False, cols(i)(3), True
        If cols(i)(4) Then
            If hasGrp Then RText tmp, 8, "=Sum([" & cols(i)(0) & "])", x, CLng(0.05 * CM), w, rowH, 9, True, cols(i)(3), True
            RText tmp, 6, "=Sum([" & cols(i)(0) & "])", x, CLng(0.15 * CM), w, rowH, 10, True, cols(i)(3), True
        ElseIf Not firstLabelDone Then
            If hasGrp Then RLabel tmp, 8, "إجمالي الموظف", x, CLng(0.05 * CM), w, rowH, 9, True, True
            RLabel tmp, 6, "الإجمالي العام", x, CLng(0.15 * CM), w, rowH, 10, True, True
            firstLabelDone = True
        End If
        x = x + w
    Next

    rpt.Section(secCols).Height = yCols + CLng(0.8 * CM)
    rpt.Section(acDetail).Height = rowH
    If hasGrp Then rpt.Section(8).Height = CLng(0.9 * CM)
    rpt.Section(6).Height = CLng(1 * CM)
    LogErr "Sections " & rName

    rpt.OnNoData = "[Event Procedure]"
    rpt.HasModule = True
    rpt.Module.AddFromString "Private Sub Report_NoData(Cancel As Integer)" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    MsgBox Forms!frmMain!lblMsgNoData.Caption, vbInformation" & vbCrLf & _
        "    Cancel = True" & vbCrLf & _
        "End Sub" & vbCrLf
    LogErr "NoData " & rName

    app.DoCmd.Close acReport, tmp, acSaveYes
    app.DoCmd.Rename rName, acReport, tmp
    LogErr "Save report " & rName
End Sub

Sub CreateReports()
    Dim period, empTitle
    period = "=""الفترة من "" & Format(" & P_FROM & ",""dd/mm/yyyy"") & ""  إلى  "" & Format(" & P_TO & ",""dd/mm/yyyy"")"
    empTitle = "=""الموظف: "" & [EmpID] & "" - "" & [EmpName] & ""     الإدارة: "" & [DeptName]"

    BuildReport "rpt01_DailyAttendance", "تقرير الحضور والانصراف اليومي التفصيلي", "qryR01_Attendance", "EmpID", empTitle, "WorkDate", Array( _
        Array("WorkDate", "التاريخ", 2.2, "dd/mm/yyyy", False), _
        Array("DayName", "اليوم", 1.8, "", False), _
        Array("TimeIn", "الحضور", 1.6, "hh:nn", False), _
        Array("TimeOut", "الانصراف", 1.6, "hh:nn", False), _
        Array("WorkedHrs", "ساعات العمل", 1.9, "0.00", True), _
        Array("LateMin", "تأخير (د)", 1.6, "0", True), _
        Array("EarlyMin", "انصراف مبكر (د)", 2, "0", True), _
        Array("OTHrs", "إضافي (س)", 1.7, "0.00", True), _
        Array("DayStatus", "الحالة", 3.6, "", False)), False, period

    BuildReport "rpt02_LateEarly", "تقرير التأخير والانصراف المبكر والبصمات الناقصة", "qryR02_LateEarly", "EmpID", empTitle, "WorkDate", Array( _
        Array("WorkDate", "التاريخ", 2.2, "dd/mm/yyyy", False), _
        Array("DayName", "اليوم", 1.8, "", False), _
        Array("TimeIn", "الحضور", 1.6, "hh:nn", False), _
        Array("TimeOut", "الانصراف", 1.6, "hh:nn", False), _
        Array("LateMin", "تأخير (د)", 1.7, "0", True), _
        Array("EarlyMin", "انصراف مبكر (د)", 2.1, "0", True), _
        Array("LateDeduct", "قيمة الخصم", 2, "#,##0.00", True), _
        Array("DayStatus", "الحالة", 4, "", False)), False, period

    BuildReport "rpt03_Absence", "تقرير الغياب بدون إذن", "qryR03_Absence", "EmpID", empTitle, "CalDate", Array( _
        Array("CalDate", "تاريخ الغياب", 4, "dd/mm/yyyy", False), _
        Array("DayName", "اليوم", 4, "", False), _
        Array("AbsCnt", "عدد الأيام", 3, "0", True)), False, period

    BuildReport "rpt04_Leaves", "تقرير الإجازات والمأموريات", "qryR04_Leaves", "EmpID", empTitle, "FromDate", Array( _
        Array("LeaveTypeName", "نوع الإجازة", 3.8, "", False), _
        Array("PaidText", "الأجر", 2, "", False), _
        Array("FromDate", "من", 2.3, "dd/mm/yyyy", False), _
        Array("ToDate", "إلى", 2.3, "dd/mm/yyyy", False), _
        Array("DaysInRange", "أيام العمل داخل الفترة", 2.8, "0", True), _
        Array("Notes", "ملاحظات", 5.5, "", False)), False, period

    BuildReport "rpt05_LeaveBalance", "تقرير أرصدة الإجازات السنوية", "qryR05_LeaveBalance", "EmpID", empTitle, "LeaveTypeID", Array( _
        Array("LeaveTypeName", "نوع الإجازة", 5, "", False), _
        Array("Entitled", "الرصيد المستحق", 3, "0", True), _
        Array("UsedD", "المستخدم", 3, "0", True), _
        Array("RemainD", "المتبقي", 3, "0", True)), False, _
        "=""عن سنة "" & Year(" & P_TO & ") & ""  (حتى تاريخه)"""

    BuildReport "rpt06_Overtime", "تقرير الساعات الإضافية وقيمتها", "qryR06_Overtime", "EmpID", empTitle, "WorkDate", Array( _
        Array("WorkDate", "التاريخ", 2.3, "dd/mm/yyyy", False), _
        Array("DayName", "اليوم", 1.9, "", False), _
        Array("TimeIn", "الحضور", 1.7, "hh:nn", False), _
        Array("TimeOut", "الانصراف", 1.7, "hh:nn", False), _
        Array("DayOTHrs", "إضافي نهاري (س)", 2.4, "0.00", True), _
        Array("NightOTHrs", "إضافي ليلي (س)", 2.4, "0.00", True), _
        Array("OffOTHrs", "راحات/عطلات (س)", 2.6, "0.00", True), _
        Array("OTHrs", "إجمالي الإضافي (س)", 2.6, "0.00", True), _
        Array("HourRate", "أجر الساعة", 2, "#,##0.00", False), _
        Array("OTPay", "قيمة الإضافي", 2.6, "#,##0.00", True), _
        Array("DayStatus", "الحالة", 3.6, "", False)), True, period

    BuildReport "rpt07_Meals", "تقرير بدلات الغداء والعشاء", "qryR07_Meals", "EmpID", empTitle, "WorkDate", Array( _
        Array("WorkDate", "التاريخ", 2.3, "dd/mm/yyyy", False), _
        Array("DayName", "اليوم", 1.9, "", False), _
        Array("TimeIn", "الحضور", 1.7, "hh:nn", False), _
        Array("TimeOut", "الانصراف", 1.7, "hh:nn", False), _
        Array("OTHrs", "إضافي (س)", 1.8, "0.00", True), _
        Array("LunchAllow", "بدل غداء", 2.2, "#,##0.00", True), _
        Array("DinnerAllow", "بدل عشاء", 2.2, "#,##0.00", True), _
        Array("MealTotal", "الإجمالي", 2.4, "#,##0.00", True)), False, period

    BuildReport "rpt08_MonthlySummary", "الملخص الشهري للحضور والانصراف", "qryR08_Summary", "", "", "EmpID", Array( _
        Array("EmpID", "الكود", 1.4, "", False), _
        Array("EmpName", "اسم الموظف", 4, "", False), _
        Array("DeptName", "الإدارة", 3, "", False), _
        Array("PresentDays", "أيام الحضور", 1.7, "0", True), _
        Array("OffWorkDays", "عمل براحة/عطلة", 1.9, "0", True), _
        Array("AbsDays", "غياب", 1.4, "0", True), _
        Array("PaidLeaveDays", "إجازات مدفوعة", 1.8, "0", True), _
        Array("UnpaidLeaveDays", "بدون أجر", 1.6, "0", True), _
        Array("LateMin", "تأخير (د)", 1.6, "0", True), _
        Array("EarlyMin", "انصراف مبكر (د)", 1.9, "0", True), _
        Array("MissingPunch", "بصمات ناقصة", 1.7, "0", True), _
        Array("WorkedHrs", "ساعات العمل", 2, "0.00", True), _
        Array("OTHrs", "ساعات إضافية", 1.9, "0.00", True)), True, period

    BuildReport "rpt09_Payroll", "كشف الأجور الأساسية والإضافية والبدلات", "qryR09_Payroll", "", "", "EmpID", Array( _
        Array("EmpID", "الكود", 1.3, "", False), _
        Array("EmpName", "اسم الموظف", 3.6, "", False), _
        Array("DeptName", "الإدارة", 2.6, "", False), _
        Array("BasicSalary", "الأجر الأساسي", 2.2, "#,##0.00", True), _
        Array("OTPay", "الأجر الإضافي", 2.1, "#,##0.00", True), _
        Array("LunchAllow", "بدل غداء", 1.7, "#,##0.00", True), _
        Array("DinnerAllow", "بدل عشاء", 1.7, "#,##0.00", True), _
        Array("Gross", "إجمالي المستحق", 2.2, "#,##0.00", True), _
        Array("AbsDeduct", "خصم الغياب", 1.9, "#,##0.00", True), _
        Array("UnpaidDeduct", "خصم بدون أجر", 1.9, "#,##0.00", True), _
        Array("LateDeduct", "خصم التأخير", 1.8, "#,##0.00", True), _
        Array("TotalDeduct", "إجمالي الخصم", 2, "#,##0.00", True), _
        Array("NetPay", "صافي المستحق", 2.3, "#,##0.00", True)), True, period

    BuildReport "rpt10_Departments", "ملخص الحضور والأجور على مستوى الإدارات", "qryR10_Dept", "", "", "DeptName", Array( _
        Array("DeptName", "الإدارة", 3.8, "", False), _
        Array("dEmpCount", "عدد الموظفين", 1.9, "0", True), _
        Array("dPresent", "أيام الحضور", 1.9, "0", True), _
        Array("dAbs", "أيام الغياب", 1.9, "0", True), _
        Array("dLeave", "أيام الإجازات", 1.9, "0", True), _
        Array("dLateHrs", "ساعات التأخير", 2, "0.00", True), _
        Array("dOTHrs", "ساعات إضافية", 2, "0.00", True), _
        Array("dOTPay", "قيمة الإضافي", 2.3, "#,##0.00", True), _
        Array("dMeals", "البدلات", 2.1, "#,##0.00", True), _
        Array("dBasic", "الأجور الأساسية", 2.4, "#,##0.00", True), _
        Array("dDeduct", "الخصومات", 2.1, "#,##0.00", True), _
        Array("dNet", "صافي المستحق", 2.5, "#,##0.00", True)), True, period
End Sub

'==========================================================================
' 7) Startup options
'==========================================================================
Sub FinalizeDb()
    On Error Resume Next
    Set db = app.CurrentDb
    SetDbProp "StartupForm", dbText, "frmMain"
    SetDbProp "AppTitle", dbText, "نظام الحضور والانصراف"
End Sub
