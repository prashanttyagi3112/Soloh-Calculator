Attribute VB_Name = "SolohCalc"
' Soloh Travel Pay Package Calculator v1.0 - Reset and Save-as-PDF macros
Option Explicit
Private Const SHEET_NAME As String = "Pay Package - Rate Calculator"
Private Const PWD As String = "soloh"

Public Sub ResetCalculator()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If MsgBox("Clear all entries and reset the calculator?", vbYesNo + vbQuestion, "Reset Calculator") <> vbYes Then Exit Sub
    ws.Unprotect PWD
    ' Client / job details
    ws.Range("C6").Value = "<Enter Client Name>"
    ws.Range("C7").Value = 0.06
    ws.Range("C10").Value = "<Enter Job Title>"
    ws.Range("C11").Value = "<99-123459>"
    ws.Range("C12").Value = "<999999999999>"
    ws.Range("D13").Value = "None"
    ' Candidate details
    ws.Range("G6").Value = "<Enter Candidate Name>"
    ws.Range("G7").Value = "<Enter Candidate Email>"
    ws.Range("G8").Value = "<Enter Recruiter Name>"
    ws.Range("G9").Value = "<Enter Recruiter Email>"
    ws.Range("G10").Value = ""
    ' Bill and pay
    ws.Range("D16").Value = 0
    ws.Range("D18").Value = 0
    ws.Range("H16").Value = "New"
    ws.Range("H17").Value = 0
    ws.Range("H19").Value = 0
    ws.Range("H20").Value = 0
    ' OT policy
    ws.Range("D23").Value = "No"
    ws.Range("D24").Value = "No OT (straight)"
    ws.Range("D25").Value = "OT after 40 hrs"
    ws.Range("H23").Value = 0
    ' Contract details
    ws.Range("D28").Value = 13
    ws.Range("D29").Value = 36
    ws.Range("D30").Value = 12
    ws.Range("D31").Value = "Soloh"
    ws.Range("D32").Value = 12
    ' Bonuses and notes
    ws.Range("H28:H32").Value = 0
    ws.Range("O33").Value = ""
    ws.Protect PWD, DrawingObjects:=False, Contents:=True, Scenarios:=True
    ws.Range("C6").Select
End Sub

Public Sub SaveOfferAsPDF()
    Dim ws As Worksheet, f As String, nm As String
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    nm = Trim(ws.Range("G6").Value)
    If nm = "" Or Left(nm, 1) = "<" Then nm = "Candidate"
    f = ThisWorkbook.Path
    If f = "" Then f = Environ("USERPROFILE") & "\Documents"
    f = f & Application.PathSeparator & "Soloh_PayPackage_" & Replace(nm, " ", "_") & "_" & Format(Now, "yyyy-mm-dd_hhmm") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "PDF saved:" & vbCrLf & f, vbInformation, "Saved"
End Sub
