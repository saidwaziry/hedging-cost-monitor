Attribute VB_Name = "HedgingMonitor"
Option Explicit

' ============================================================
' Hedging-Kosten-Monitor: Datenupdate aus FRED
' Laedt vier monatliche Zinsreihen (OECD via FRED) und fuellt
' das Blatt "Historie". Formeln in F bis J werden pro Zeile gesetzt.
' ============================================================

Private Const FRED_URL As String = _
    "https://fred.stlouisfed.org/graph/fredgraph.csv?id=" & _
    "IR3TIB01USM156N,IR3TIB01EZM156N,IRLTLT01USM156N,IRLTLT01DEM156N&cosd=2000-01-01"

Private Const ERSTE_ZEILE As Long = 5
Private Const MAX_ZEILEN As Long = 500

Public Sub DatenAktualisieren()
    Dim http As Object, txt As String
    Dim zeilen() As String, felder() As String
    Dim ws As Worksheet
    Dim i As Long, r As Long, k As Long
    Dim werte(1 To 4) As Double, vollstaendig As Boolean

    Set ws = ThisWorkbook.Worksheets("Historie")

    On Error GoTo Fehler
    Application.StatusBar = "Lade Daten von FRED ..."
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", FRED_URL, False
    http.send
    If http.Status <> 200 Then
        MsgBox "FRED antwortet mit Status " & http.Status & ". Bitte spaeter erneut versuchen.", vbExclamation
        GoTo Ende
    End If
    txt = Replace(http.responseText, vbCr, "")
    zeilen = Split(txt, vbLf)

    Application.ScreenUpdating = False
    ws.Range(ws.Cells(ERSTE_ZEILE, 1), ws.Cells(ERSTE_ZEILE + MAX_ZEILEN - 1, 10)).ClearContents

    r = ERSTE_ZEILE
    For i = 1 To UBound(zeilen)              ' Zeile 0 ist die Kopfzeile
        If Len(Trim(zeilen(i))) > 0 Then
            felder = Split(zeilen(i), ",")
            If UBound(felder) >= 4 Then
                vollstaendig = True
                For k = 1 To 4
                    If felder(k) = "" Or felder(k) = "." Then
                        vollstaendig = False
                    Else
                        werte(k) = Val(felder(k)) / 100   ' Val nutzt immer den Punkt als Dezimalzeichen
                    End If
                Next k

                ' Nur Monate mit allen vier Werten uebernehmen (keine Luecken im Chart)
                If vollstaendig And r < ERSTE_ZEILE + MAX_ZEILEN Then
                    ws.Cells(r, 1).Value = DateSerial(CInt(Left(felder(0), 4)), CInt(Mid(felder(0), 6, 2)), 1)
                    For k = 1 To 4
                        ws.Cells(r, 1 + k).Value = werte(k)
                    Next k
                    ws.Cells(r, 6).Formula = "=((1+B" & r & "*Rechner!$B$6/Rechner!$B$7)/(1+C" & r & _
                                             "*Rechner!$B$6/Rechner!$B$7)-1)*Rechner!$B$7/Rechner!$B$6"
                    ws.Cells(r, 7).Formula = "=D" & r & "-F" & r
                    ws.Cells(r, 8).Formula = "=(G" & r & "-E" & r & ")*10000"
                    ws.Cells(r, 9).Formula = "=D" & r & "-B" & r
                    ws.Cells(r, 10).Formula = "=E" & r & "-C" & r
                    r = r + 1
                End If
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    If r = ERSTE_ZEILE Then
        MsgBox "Keine vollstaendigen Daten gefunden.", vbExclamation
    Else
        MsgBox (r - ERSTE_ZEILE) & " Monate geladen. Letzter vollstaendiger Monat: " & _
               Format(ws.Cells(r - 1, 1).Value, "mm.yyyy"), vbInformation
    End If
    GoTo Ende

Fehler:
    Application.ScreenUpdating = True
    MsgBox "Fehler beim Laden: " & Err.Description & vbNewLine & _
           "Pruefe die Internetverbindung (Windows-Excel erforderlich).", vbCritical
Ende:
    Application.StatusBar = False
End Sub

Public Sub NeuesterStandInRechner()
    Dim ws As Worksheet, rc As Worksheet, letzte As Long

    Set ws = ThisWorkbook.Worksheets("Historie")
    Set rc = ThisWorkbook.Worksheets("Rechner")
    letzte = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If letzte < ERSTE_ZEILE Then
        MsgBox "Noch keine Daten. Bitte zuerst 'DatenAktualisieren' ausfuehren.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Die Werte aus " & Format(ws.Cells(letzte, 1).Value, "mm.yyyy") & _
              " ueberschreiben Rechner B8 bis B11. Fortfahren?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    rc.Range("B8").Value = ws.Cells(letzte, 2).Value   ' USD 3M
    rc.Range("B9").Value = ws.Cells(letzte, 3).Value   ' EUR 3M
    rc.Range("B10").Value = ws.Cells(letzte, 4).Value  ' UST 10J
    rc.Range("B11").Value = ws.Cells(letzte, 5).Value  ' Bund 10J
    MsgBox "Rechner aktualisiert (Monatsdurchschnitte, nicht Tageswerte).", vbInformation
End Sub
