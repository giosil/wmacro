Public Type Invoice
    
    Type As String
    Number As String
    SentDate As Date
    DocNumber As String
    DocDate As Date
    Amount As Double
    TaxRate As Double
    Total As Double

End Type

Public Type InvoiceTotals
    
    Amount As Double
    VAT As Double
    Total As Double

End Type

Sub LoadInvoices()

Dim dataFile As String, textline As String, row As Integer, totRow As Boolean, vatRow As Boolean, inv As Invoice, tot As InvoiceTotals

dataFile = Application.GetOpenFilename()

If dataFile = "" Or dataFile = "Falso" Or dataFile = "False" Then
    
    Exit Sub

End If

Call InitSheet

Open dataFile For Input As #1

row = 1
totRow = False
vatRow = False

Do Until EOF(1)
    Line Input #1, textline
    
    If textline Like "F*" Then
        inv = ParseInvoice(textline)
        If inv.Type <> "" Then
            row = row + 1
                    
            Range("A" & row).NumberFormat = "@"
            Range("A" & row).Value = inv.Type
            
            Range("B" & row).NumberFormat = "@"
            Range("B" & row).Value = inv.Number
            
            Range("C" & row).NumberFormat = "dd/mm/yyyy"
            Range("C" & row).Value = inv.SentDate
            
            Range("D" & row).NumberFormat = "@"
            Range("D" & row).Value = inv.DocNumber
            
            Range("E" & row).NumberFormat = "dd/mm/yyyy"
            Range("E" & row).Value = inv.DocDate
            
            totRow = True
            vatRow = False
        End If
    End If
    
    If InStr(textline, "Totali") > 0 And totRow Then
        tot = ParseTotals(textline)
    
        Range("F" & row).Value = tot.Amount
        Range("G" & row).Value = tot.VAT
        Range("H" & row).Value = tot.Total
        Range("I" & row).Value = 0
                
        totRow = False
        vatRow = True
    End If
    
    If InStr(textline, "%") > 0 And vatRow Then
        Range("I" & row).Value = GetTaxRate(textline)
        
        totRow = False
        vatRow = False
    End If
    
Loop

Close #1

Cells.Select
Cells.EntireColumn.AutoFit
Range("A1").Select

End Sub

Private Sub InitSheet()

ActiveSheet.Select
ActiveSheet.Cells.Select
Selection.ClearContents

Range("A1").NumberFormat = "@"
Range("A1").Font.Bold = True
Range("A1").Value = "Tipo"

Range("B1").NumberFormat = "@"
Range("B1").Font.Bold = True
Range("B1").Value = "Numero"

Range("C1").NumberFormat = "@"
Range("C1").Font.Bold = True
Range("C1").Value = "Data"

Range("D1").NumberFormat = "@"
Range("D1").Font.Bold = True
Range("D1").Value = "Documento"

Range("E1").NumberFormat = "@"
Range("E1").Font.Bold = True
Range("E1").Value = "Data Doc."

Range("F1").NumberFormat = "@"
Range("F1").Font.Bold = True
Range("F1").Value = "Imponibile"

Range("G1").NumberFormat = "@"
Range("G1").Font.Bold = True
Range("G1").Value = "IVA"

Range("H1").NumberFormat = "@"
Range("H1").Font.Bold = True
Range("H1").Value = "Totale"

Range("I1").NumberFormat = "@"
Range("I1").Font.Bold = True
Range("I1").Value = "Aliquota IVA"

End Sub

Private Function ParseInvoice(textline As String) As Invoice

Dim result As Invoice, s0 As Integer, s1 As Integer, s2 As Integer, s3 As Integer, s4 As Integer

result.Type = ""

s0 = InStr(textline, " ")
s1 = InStr(textline, "del")
s2 = InStr(s1, textline, "Documento")
s3 = InStr(s2, textline, "del")
s4 = InStr(s3, textline, "Comp")

If s0 > 0 And s1 > s0 And s2 > s1 And s3 > s2 And s4 > s3 Then
    result.Type = Trim(Mid(textline, 1, s0 - 1))
    result.Number = Trim(Mid(textline, s0, s1 - s0))
    result.SentDate = toDate(Mid(textline, s1 + 3, s2 - s1 - 3))
    result.DocNumber = Trim(Mid(textline, s2 + 9, s3 - s2 - 9))
    result.DocDate = toDate(Mid(textline, s3 + 3, s4 - s3 - 3))
End If

ParseInvoice = result

End Function

Private Function ParseTotals(textline As String) As InvoiceTotals

Dim result As InvoiceTotals, i As Integer, a As Double, v As Double, cols() As String

cols = Split(textline, "|")

For i = 1 To UBound(cols)
    
    Select Case i
        Case 2
            result.Amount = toDouble(cols(i))
        Case 3
            result.VAT = toDouble(cols(i))
        Case 5
            result.Total = toDouble(cols(i))
    End Select

Next

ParseTotals = result

End Function

Private Function GetTaxRate(textline As String) As Double

Dim result As Double, b As Integer, e As Double

result = 0

e = InStr(textline, "%")

If e > 0 Then

    b = InStr(e - 3, textline, " ")
    
    If b > 0 Then
        
        result = toDouble(Mid(textline, b, e - b))
    
    End If

End If

GetTaxRate = result

End Function

Private Function toDouble(v As String) As Double

Dim result As Double

On Error GoTo LabelErr

result = CDbl(Trim(Replace(v, ".", ",")))

toDouble = result

Exit Function

LabelErr:
    toDouble = 0

End Function

Private Function toDate(v As String) As Date

Dim result As Date, s1 As Integer, s2 As Integer, y As Integer, m As Integer, d As Integer

On Error GoTo LabelErr

v = Trim(v)

s1 = InStr(v, "/")

If s1 > 0 Then

    s2 = InStr(s1 + 1, v, "/")
    
    If s2 > 0 Then
        
        d = CInt(Trim(Mid(v, 1, s1 - 1)))
        m = CInt(Trim(Mid(v, s1 + 1, s2 - s1 - 1)))
        y = CInt(Trim(Mid(v, s2 + 1)))
        
        If y < 1000 Then
            y = 2000 + y
        End If
        
        result = DateSerial(y, m, d)
    
    End If

End If

toDate = result

Exit Function

LabelErr:

End Function

