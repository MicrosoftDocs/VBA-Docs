---
title: Series object (Project)
ms.prod: project-server
ms.assetid: 38a834ec-4076-82ef-a6bd-55a1ee2624bd
ms.date: 06/08/2017
localization_priority: Normal
---


# Series object (Project)
Represents a collection of related data that makes a row or a column in a chart.
 

## Remarks

A **Series** object is a member of the **[SeriesCollection](Project.seriescollection.md)** collection, which includes all of the data series in the chart. The name of the series is often displayed in the chart legend.
 

 

## Example

The following example prints the series names, X (horizontal) values, and Y (vertical) values for a collection of data series on a chart.
 

 

```vb
Sub TestChartSeries()
    Dim reportName As String
    Dim theReportIndex As Integer
    Dim theChart As Chart
    Dim seriesCollec As SeriesCollection
    Dim chartSeries As Series
    Dim i As Integer
    Dim j As Integer
        
    reportName = "Simple scalar chart"
    theReportIndex = -1
        
    If (ActiveProject.Reports.IsPresent(reportName)) Then
        ' Make the report active.
        theReportIndex = ActiveProject.Reports(reportName).Index
        ActiveProject.Reports(theReportIndex).Apply
        
        Set theChart = ActiveProject.Reports(theReportIndex).Shapes(1).Chart
        Set seriesCollec = theChart.SeriesCollection()
        
        For i = 1 To seriesCollec.Count
            Set chartSeries = seriesCollec(i)
        
            If (IsEmpty(chartSeries.Name)) Then
                Debug.Print "Series " & i & " name is an empty string."
            Else
                Debug.Print "Series " & i & ": " & chartSeries.Name
            End If
            
            For j = 1 To seriesCollec.Count
                Debug.Print vbTab & "X, Y values(" & j & "): " & chartSeries.XValues(j) _
                    & ", " & chartSeries.Values(j); ""
            Next j
        Next i
    End If
End Sub
```

The following sample output is from a chart such as the example in the [Chart](Project.chart.md) object documentation.
 

 



```vb
Series 1: Actual Work
    X, Y values(1): T1, 16
    X, Y values(2): T2 - new, 32
    X, Y values(3): T3, 7
Series 2: Remaining Work
    X, Y values(1): T1, 0
    X, Y values(2): T2 - new, 16
    X, Y values(3): T3, 17
Series 3: Work
    X, Y values(1): T1, 16
    X, Y values(2): T2 - new, 48
    X, Y values(3): T3, 24
```


## Properties



|Name|
|:-----|
|[Application](Project.series.application.md)|
|[Name](Project.series.name.md)|
|[Parent](Project.series.parent.md)|
|[Values](Project.series.values.md)|
|[XValues](Project.series.xvalues.md)|

## See also


 
[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]