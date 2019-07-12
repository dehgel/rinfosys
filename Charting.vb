Module Charting
    Function _GetPopulation() As Single
        Dim r = Rinfosys.Chart1
        ' Add data points to the two series
        Dim random As New Random()
        Dim pointIndex As Integer
        For pointIndex = 0 To 9
            r.Series("Population").Points.AddY(random.Next(45, 95))

            r.Series("Residents").Points.AddY(random.Next(5, 75))
        Next pointIndex
    End Function
End Module
