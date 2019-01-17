---
title: Working with Charts
ms.prod: word
ms.assetid: 7afe145a-f8fb-0123-c105-de1dde11db9e
ms.date: 06/08/2017
localization_priority: Normal
---


# Working with Charts

In Word 2007 Service Pack 2 (SP2) and later, you can programmatically access and manipulate charts using the VBA object model in Word. The chart object in Word is drawn by the same shared Office drawing layer implementation used by Excel, so if you are familiar with the charting object model in Excel, you can easily migrate Excel VBA code that manipulates charts into Word VBA code.


## Using the Chart Object

In Word, a chart is represented by a  [Chart](../../../api/Word.Chart.md) object. The [Chart](../../../api/Word.Chart.md) object is contained by an InlineShape or Shape. You can use either the [InlineShapes](../../../api/Word.shapes.md) collection or the [Shapes](../../../api/Word.shapes.md) collection of the [Document](../../../api/Word.Document.md) object to add new or access existing charts. You use the [AddChart](../../../api/overview/Word.md) method for both collections, specifying the chart type and location within the document, to add a new chart.

You can use the  [HasChart](../../../api/Word.InlineShape.HasChart.md) property to determine if an [InlineShape](../../../api/Word.Shape.md) object or [Shape](../../../api/Word.Shape.md) object contains a chart. If [HasChart](../../../api/Word.InlineShape.HasChart.md) returns True, you can then use the [Chart](../../../api/Word.InlineShape.Chart.md) property to get a reference to a [Chart](../../../api/Word.Chart.md) object that represents the chart. At this point, the implementation is virtually identical as that of Excel and VBA code can be transferred between the two programs in most cases.

For example, the following VBA code example adds a new 2-D stacked column chart to the active worksheet in Excel and sets the chart's source data to the range A1:C3 from the Sheet1 worksheet.




```vb
Sub AddChart_Excel()
    Dim objShape As Shape
    
    ' Create a chart and return a Shape object reference.
    ' The Shape object reference contains the chart.
    Set objShape = ActiveSheet.Shapes.AddChart(XlChartType.xlColumnStacked100)
    
    ' Ensure the Shape object contains a chart. If so,
    ' set the source data for the chart to the range A1:C3.
    If objShape.HasChart Then
        objShape.Chart.SetSourceData Source:=Range("'Sheet1'!$A$1:$C$3")
    End If
End Sub
```

By comparison, the following VBA code example adds a new 2-D stacked column chart to the active document in and sets the chart's source data to the range A1:C3 from the chart data associated with the chart. 




```vb
Sub AddChart_Word()
    Dim objShape As InlineShape
    
    ' Create a chart and return a Shape object reference.
    ' The Shape object reference contains the chart.
    Set objShape = ActiveDocument.InlineShapes.AddChart(XlChartType.xlColumnStacked100)
    
    ' Ensure the Shape object contains a chart. If so,
    ' set the source data for the chart to the range A1:C3.
    If objShape.HasChart Then
        objShape.Chart.SetSourceData Source:="'Sheet1'!$A$1:$C$3"
    End If
End Sub
```


## Key Differences Between the Chart object in Word and the ChartObject object in Excel

Even though how you work with charts between Excel and Word is nearly identical in most cases, it is helpful to identify important areas where the two implementations differ: 




- Programmatically creating or manipulating a  [ChartData](../../../api/Word.ChartData.md) object in Word requires Excel to run.
    
- Chart properties and methods for manipulating the chart sheet are not implemented. The concept of a chart sheet is specific to Excel. Chart sheets are not used in Word, so methods and properties used to reference or manipulate a chart sheet have been disabled for those applications.
    
- Properties and methods that, in Excel normally take a  [Range](../../../api/Excel.Range(object).md) object reference now take a range address in Word. The [Range](../../../api/Word.Range.md) object in Word is different than the [Range](../../../api/Excel.Range(object).md) object in Excel. To prevent confusion, the charting object model in Word accepts range address strings, such as "='Sheet1'!$A$1:$D$5", in those properties and methods (such as the [SetSourceData](../../../api/Word.Chart.SetSourceData.md) method of the Chart object) that accept Range objects in Excel.
    
- A new object,  [ChartData](../../../api/Word.ChartData.md), has been added to the VBA object models for Word to provide access to the underlying linked or embedded data for a chart. Each chart has, associated with it, the data used to draw the chart in Word. The chart data can either be linked from an external Excel workbook, or embedded as part of the chart itself. The  [ChartData](../../../api/Word.ChartData.md) object encapsulates access to the data for a given chart in Word. For example, the following VBA code example displays, and then minimizes, the chart data for each chart contained by the active document in Word.
    



```vb
Sub ShowWorkbook_Word() 
    Dim objShape As InlineShape 
     
    ' Iterates each inline shape in the active document. 
    ' If the inline shape contains a chart, then display the 
    ' data associated with that chart and minimize the application 
    ' used to display the data. 
    For Each objShape In ActiveDocument.InlineShapes 
        If objShape.HasChart Then 
 
            ' Activate the topmost window of the application used to 
            ' display the data for the chart. 
            objShape.Chart.ChartData.Activate 
             
            ' Minimize the application used to display the data for 
            ' the chart. 
            objShape.Chart.ChartData.Workbook.Application.WindowState = -4140 
        End If 
    Next 
End Sub 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]