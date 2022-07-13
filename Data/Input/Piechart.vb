Sub Consolidate()
    ActiveSheet.Shapes.AddChart2(262, xl3DPie).Select
    ActiveChart.SetSourceData Source:=Range("Consolidate_Sheet!$A$1:$H$6")
    ActiveChart.ChartTitle.Select
    Selection.Left = 80.097
    Selection.Top = 0
    ActiveChart.PlotArea.Select
    Selection.Left = 22
    Selection.Top = 41.18
    ActiveChart.ChartArea.Select
End Sub