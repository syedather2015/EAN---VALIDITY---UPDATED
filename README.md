# 🧾 EAN Validator & Highlighter – Excel VBA Tool

This Excel macro streamlines the process of validating **EAN (European Article Numbers)** by checking formatting rules, ensuring correct length, and automatically applying **13-digit padding**. Invalid entries are visually highlighted, making it easier to spot and correct them during product data onboarding or audit processes.

> Developed by **Syed Ather Rizvi** for product data integrity and quality assurance.

---

## ⚙️ Features

- ✅ Prompts user to select the column with EANs
- 🔢 Automatically pads EANs with leading zeros to make them 13 digits
- 🧠 Applies rule-based checks (no leading "2", no "00000" patterns, etc.)
- 🧮 **Includes EAN-13 check digit validation**
- 🎨 Highlights invalid EANs with **blue background** and **white text**
- 📊 Works on the **active worksheet** – no setup required
- 🔤 Supports column letters up to **XFD** (e.g., A, Z, AA, AB...)

---

## 🧪 Validation Rules Applied

A valid EAN must:
- Be **numeric**
- Be **exactly 13 digits** (auto-padded if shorter)
- **Not start with `2`**
- **Not contain `00000`** in:
  - Characters 1–3
  - Characters 3–7
  - Characters 8–12
- Pass the official **EAN-13 checksum digit** validation

---

## 🔁 How to Use

1. Open your Excel file with EAN data
2. Press `Alt + F11` to open the **VBA Editor**
3. Insert a new **Module** and paste the macro code
4. Close the editor and run `ValidateAndHighlight` from the Macro window (`Alt + F8`)
5. When prompted, enter the **column letter** where EANs exist (e.g., `B`, `D`, `AA`, etc.)

---

## 💡 Use Cases

- Product onboarding & GTIN validation  
- Retail or eCommerce data audits  
- Supplier catalog quality checks  
- Marketing intelligence & pricing validation  

---

## 📌 Visual Output

- Valid EANs:
  - Auto-corrected to 13 digits if shorter
- Invalid EANs:
  - Highlighted with a **blue fill** and **white font** for visibility

---

## 📄 Code Highlights

```vba
Sub ValidateAndHighlight()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim eanRange As Range
    Dim cell As Range
    Dim columnLetter As String

    columnLetter = InputBox("Enter the column letter containing EAN codes (e.g., A, B, C)", "Column Selection")

    If Not IsValidColumn(columnLetter) Then
        MsgBox "Invalid column letter. Exiting macro."
        Exit Sub
    End If

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
    Set eanRange = ws.Range(columnLetter & "2:" & columnLetter & lastRow)

    For Each cell In eanRange
        If Not IsEmpty(cell) Then
            Dim ean As String
            ean = CStr(cell.Value)
            If Len(ean) < 13 Then
                ean = WorksheetFunction.Rept("0", 13 - Len(ean)) & ean
            End If
            If IsValidEAN(ean) Then
                cell.Value = ean
            Else
                HighlightCell cell
            End If
        End If
    Next cell
End Sub

Function IsValidEAN(ByVal ean As String) As Boolean
    If Len(ean) <> 13 Or Not IsNumeric(ean) Then Exit Function
    If Left(ean, 1) = "2" Then Exit Function
    If Mid(ean, 1, 3) = "000" Or Mid(ean, 3, 5) = "00000" Or Mid(ean, 8, 5) = "00000" Then Exit Function
    IsValidEAN = IsValidEAN13CheckDigit(ean)
End Function

Function IsValidEAN13CheckDigit(ean As String) As Boolean
    Dim i As Integer, sum As Long, checkDigit As Integer
    sum = 0
    For i = 1 To 12
        If i Mod 2 = 0 Then
            sum = sum + CInt(Mid(ean, i, 1)) * 3
        Else
            sum = sum + CInt(Mid(ean, i, 1))
        End If
    Next i
    checkDigit = (10 - (sum Mod 10)) Mod 10
    IsValidEAN13CheckDigit = (checkDigit = CInt(Right(ean, 1)))
End Function

Function IsValidColumn(ByVal columnLetter As String) As Boolean
    On Error Resume Next
    IsValidColumn = Not IsEmpty(Range(columnLetter & "1"))
    On Error GoTo 0
End Function

Sub HighlightCell(ByRef cell As Range)
    With cell.Interior
        .Pattern = xlSolid
        .Color = RGB(0, 176, 240)
    End With
    With cell.Font
        .Color = RGB(255, 255, 255)
    End With
End Sub
