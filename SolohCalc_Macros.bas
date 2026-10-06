Attribute VB_Name = "SolohCalc"
' Soloh Travel Pay Package Calculator v1.1 - Reset and Save-as-PDF macros
Option Explicit
Private Const SHEET_NAME As String = "Pay Package - Rate Calculator"
Private Const PWD As String = "soloh"

Public Sub ResetCalculator()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If MsgBox("Clear all entries and reset the calculator?", vbYesNo + vbQuestion, "Reset Calculator") <> vbYes Then Exit Sub
    ws.Unprotect PWD
    ' Client details
    ws.Range("C6").Value = "<Enter Client Name>"
    ws.Range("C7").Value = 0.06
    ws.Range("D8").Value = "None"
    ' Bill and pay
    ws.Range("D11").Value = 0
    ws.Range("D13").Value = 0
    ws.Range("H6").Value = "New"
    ws.Range("H7").Value = 0
    ws.Range("H9").Value = 0
    ws.Range("H10").Value = 0
    ' OT policy
    ws.Range("D16").Value = "No"
    ws.Range("D17").Value = "No OT (straight)"
    ws.Range("D18").Value = "OT after 40 hrs"
    ws.Range("H13").Value = 0
    ' Contract details
    ws.Range("D21").Value = 13
    ws.Range("D22").Value = 36
    ws.Range("D23").Value = 12
    ws.Range("D24").Value = "Soloh"
    ws.Range("D25").Value = 12
    ' Bonus / reimbursement and notes
    ws.Range("H17").Value = 0
    ws.Range("O27").Value = ""
    ws.Range("L32").Value = 0   ' Target Net Margin $ off
    ws.Protect PWD, DrawingObjects:=False, Contents:=True, Scenarios:=True
    ws.Range("C6").Select
End Sub

Public Sub SaveOfferAsPDF()
    Dim ws As Worksheet, f As String, nm As String
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    nm = Trim(ws.Range("C6").Value)
    If nm = "" Or Left(nm, 1) = "<" Then nm = "Client"
    f = ThisWorkbook.Path
    If f = "" Then f = Environ("USERPROFILE") & "\Documents"
    f = f & Application.PathSeparator & "Soloh_PayPackage_" & Replace(nm, " ", "_") & "_" & Format(Now, "yyyy-mm-dd_hhmm") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "PDF saved:" & vbCrLf & f, vbInformation, "Saved"
End Sub
