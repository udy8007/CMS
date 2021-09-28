Attribute VB_Name = "Module1"
Public DESIG As String
Public NewRecord As Boolean
Public mCourceCode As String
Public mTeacherCode As String
Public mSubjectCode As String
Public md As Long
Public mUnits As String
Public mExamDate As Date
Public mYear As String
Public MySQL As String
Public CCode As String
Public year As Integer
Public attendancedate As Date
Public tcode As Integer
Public scode As String
'Public CBOAREA As String
Public adodc As ADODB.connection
Public fac As String
Public Sub connection()
Set adodc = New ADODB.connection
adodc.Open "Provider = sqloledb;" & _
          "Data Source=UDY8007\UDYSERVER;" & _
          "Initial Catalog=SDC;" & _
          "User ID=SDCUser;" & _
          "Password=SDCpwd;"
End Sub

Public Sub close_form()
Unload changepassfrm
Unload adduserfrm
Unload deluserfrm
Unload query
Unload Master_Course
Unload Master_Holiday
Unload master_item
Unload Master_staff
Unload Master_Subject
Unload Master_vendor
Unload frmAttendance
Unload frmMarksEntry
Unload Report1
Unload Report2
Unload Report3
Unload frmStaffAttendance
Unload frmVendorItems
Unload fyaddmission

Unload query
Unload staffreport

End Sub
