Attribute VB_Name = "Modulo3"
Function Classi(dati As Range) As Variant
    Dim N As Long, k As Long
    Dim MinValue As Double, MaxValue As Double
    Dim Intervallo As Double
    Dim media As Double, scarto As Double, varianza As Double
    Dim i As Long, j As Long
    Dim da() As Double, a() As Double
    Dim numerosita() As Long, percentuale() As Double
    Dim result() As Variant

    ' Calcolo del numero di dati
    N = dati.Count
    If N = 0 Then
        Classi = "Nessun dato"
        Exit Function
    End If

    ' Calcolo di k e delle statistiche base
    k = WorksheetFunction.RoundUp(1 + WorksheetFunction.Log10(N) / WorksheetFunction.Log10(2), 0)
    ' equivalente alla più nota: k = WorksheetFunction.RoundUp(1 + 3.322 * WorksheetFunction.Log10(N), 0)
    MinValue = WorksheetFunction.Min(dati)
    MaxValue = WorksheetFunction.Max(dati)
    media = WorksheetFunction.Average(dati)
    scarto = WorksheetFunction.StDev(dati)
    varianza = WorksheetFunction.VarP(dati)
    Intervallo = (MaxValue - MinValue) / k

    ' Inizializzazione degli array
    ReDim da(1 To k), a(1 To k), numerosita(1 To k), percentuale(1 To k)
    ReDim result(1 To k + 6, 1 To 5)

    ' Calcolo intervalli
    For i = 1 To k
        da(i) = MinValue + (i - 1) * Intervallo
        a(i) = MinValue + i * Intervallo
    Next i

    ' Conteggio dei dati per intervalli
    For Each cell In dati
        If IsNumeric(cell.Value) Then
            For j = 1 To k
                If (cell.Value >= da(j) And cell.Value < a(j)) Or (j = k And cell.Value = a(j)) Then
                    numerosita(j) = numerosita(j) + 1
                    Exit For
                End If
            Next j
        End If
    Next cell

    ' Preparazione dell'array dei risultati
    result(1, 1) = "Classi k": result(1, 2) = "Da": result(1, 3) = "A"
    result(1, 4) = "Numerosità": result(1, 5) = "Percentuale"

    For j = 1 To k
        result(j + 1, 1) = j
        result(j + 1, 2) = da(j)
        result(j + 1, 3) = a(j)
        result(j + 1, 4) = IIf(numerosita(j) > 0, numerosita(j), "")
        percentuale(j) = numerosita(j) / N
        result(j + 1, 5) = IIf(percentuale(j) > 0, percentuale(j), "")
    Next j

    ' Aggiunta delle statistiche
    Dim offset As Long: offset = k + 2
    result(offset, 1) = "Statistiche:"
    result(offset + 1, 1) = "Numerosità del campione": result(offset + 1, 2) = N
    result(offset + 1, 4) = "Scostamento": result(offset + 1, 5) = scarto
    result(offset + 2, 1) = "Minimo": result(offset + 2, 2) = MinValue
    result(offset + 2, 4) = "Varianza": result(offset + 2, 5) = varianza
    result(offset + 3, 1) = "Massimo": result(offset + 3, 2) = MaxValue
    result(offset + 4, 1) = "Media": result(offset + 4, 2) = media

Dim idx, idy As Integer
    For idx = k + 5 To k + 6
    For idy = 3 To 5
        result(idx, idy) = ""
    Next idy
Next idx
For idx = k + 2 To k + 2
    For idy = 2 To 5
        result(idx, idy) = ""
    Next idy
Next idx
For idx = k + 2 To k + 6
    For idy = 3 To 3
        result(idx, idy) = ""
    Next idy
Next idx

   
    ' Restituzione della matrice dei risultati
    Classi = result
End Function
