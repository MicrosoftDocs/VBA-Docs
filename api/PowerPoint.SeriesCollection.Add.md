---
title: SeriesCollection.Add method (PowerPoint)
keywords: vbapp10.chm717002
f1_keywords:
- vbapp10.chm717002
ms.prod: powerpoint
api_name:
- PowerPoint.SeriesCollection.Add
ms.assetid: 29dd05a7-a707-78ff-fc06-1085e065eb3c
ms.date: 06/08/2017
localization_priority: Normal
---


# SeriesCollection.Add method (PowerPoint)

Adds one or more new series to the collection.


## Syntax

_expression_.**Add** (_Source_, _Rowcol_, _SeriesLabels_, _CategoryLabels_, _Replace_)

_expression_ A variable that represents a '[SeriesCollection](PowerPoint.SeriesCollection.md)' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**Variant**|The new data as a string representation of a range contained in the  **[Workbook](PowerPoint.ChartData.Workbook.md)** property of the **[ChartData](PowerPoint.ChartData.md)** object for the chart.|
| _Rowcol_|Optional|**[XlRowCol](PowerPoint.XlRowCol.md)**|One of the enumeration values that specifies whether the new values are in the rows or columns of the specified range.|
| _SeriesLabels_|Optional|**Variant**|**True** if the first row or column contains the name of the data series. **False** if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Word attempts to determine the location of the series name from the contents of the first row or column.|
| _CategoryLabels_|Optional|**Variant**|**True** if the first row or column contains the name of the category labels. **False** if the first row or column contains the first data point of the series. If this argument is omitted, Word attempts to determine the location of the category label from the contents of the first row or column.|
| _Replace_|Optional|**Variant**|If CategoryLabels is  **True** and Replace is **True**, the specified categories replace the categories that currently exist for the series. If Replace is **False**, the existing categories will not be replaced. The default is **False**.|

## Return value

A  **[Series](PowerPoint.Series.md)** object that represents the new series.


## Remarks

This method does not actually return a  **Series** object as stated in the Object Browser.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example creates a new series for the first chart in the active document. The data source for the new series is range  `B1:B10` on the workbook associated with the chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection.Add _
            Source:="Sheet1!B1:B10"
    End If
End With
```


## See also


[SeriesCollection Object](PowerPoint.SeriesCollection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]