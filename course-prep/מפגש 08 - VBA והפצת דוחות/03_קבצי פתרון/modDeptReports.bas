Attribute VB_Name = "modDeptReports"
' =====================================================================
' מפגש 8 | הפצת דוח אישי לכל מנהל מחלקה, ב-PDF ובמייל
' קובץ: נובה - הפצת דוחות למנהלים.xlsm
'
' הקוד לא מכיל שמות של גיליונות או טקסט בעברית.
' הכל מגיע מהקובץ, דרך שמות מוגדרים וטבלאות:
'   DeptCell      התא שבו בוחרים מחלקה בגיליון הדוח
'   ReportYear    שנה,  ReportMonth  עד חודש
'   OutputFolder  תיקייה ל-PDF (ריק = התיקייה של הקובץ)
'   SendNow       TRUE = לשלוח מיד, FALSE = לפתוח טיוטה בלבד
'   MailSubject, MailBody   נושא וגוף המייל (נוסחאות בגיליון ההגדרות)
'   tblManagers   מחלקה | שם מנהל | מייל | לשלוח
' =====================================================================
Option Explicit

Public Sub SendDeptReports()
    Dim wsReport As Worksheet
    Dim loManagers As ListObject
    Dim row As ListRow
    Dim folder As String, pdfPath As String
    Dim olApp As Object, olMail As Object
    Dim deptNo As Variant, originalDept As Variant
    Dim countDone As Long

    Set wsReport = ThisWorkbook.Names("DeptCell").RefersToRange.Worksheet
    Set loManagers = GetTable("tblManagers")
    If loManagers Is Nothing Then
        MsgBox "Table tblManagers not found", vbExclamation
        Exit Sub
    End If

    ' תיקייה: מההגדרות, או התיקייה של הקובץ. תמיד עם \ בסוף
    folder = Trim$(CStr(Range("OutputFolder").Value))
    If folder = "" Then folder = ThisWorkbook.Path
    If Right$(folder, 1) <> "\" Then folder = folder & "\"

    originalDept = Range("DeptCell").Value
    Set olApp = CreateObject("Outlook.Application")

    For Each row In loManagers.ListRows
        ' עמודה 4: לשלוח (TRUE/FALSE)
        If row.Range.Cells(1, 4).Value = True Then
            deptNo = row.Range.Cells(1, 1).Value

            ' 1. מחליפים מחלקה בדוח ומחשבים מחדש
            Range("DeptCell").Value = deptNo
            Application.Calculate

            ' 2. שומרים PDF
            pdfPath = folder & "Report_" & deptNo & "_" & _
                      Range("ReportYear").Value & "-" & Format(Range("ReportMonth").Value, "00") & ".pdf"
            wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
                Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False

            ' 3. מייל עם הקובץ המצורף
            Set olMail = olApp.CreateItem(0)
            With olMail
                .To = row.Range.Cells(1, 3).Value          ' עמודה 3: מייל
                .Subject = Range("MailSubject").Value
                .HTMLBody = "<div dir=""rtl"" style=""font-family:Arial;font-size:11pt"">" & _
                            Replace(Range("MailBody").Value, vbLf, "<br>") & "</div>"
                .Attachments.Add pdfPath
                If Range("SendNow").Value = True Then
                    .Send
                Else
                    .Display
                End If
            End With
            countDone = countDone + 1
        End If
    Next row

    ' מחזירים את הדוח למחלקה שהייתה בו
    Range("DeptCell").Value = originalDept
    Application.Calculate

    MsgBox countDone & " reports created in " & folder, vbInformation
End Sub

' מחפש טבלה לפי שם בכל הגיליונות, בלי לכתוב שם גיליון בקוד
Private Function GetTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = tableName Then
                Set GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function
