---
title: SeriesCollection.Add method (Word)
keywords: vbawd10.chm150405301
f1_keywords:
- vbawd10.chm150405301
ms.prod: word
api_name:
- Word.SeriesCollection.Add
ms.assetid: 26778898-aa61-54f9-4db2-d38ab1399405
ms.date: 06/08/2017
localization_priority: Normal
---


# SeriesCollection.Add method (Word)

Adds one or more new series to the collection.


## Syntax

_expression_.**Add** (_Source_, _Rowcol_, _SeriesLabels_, _CategoryLabels_, _Replace_)

_expression_ A variable that represents a **[SeriesCollection](Word.SeriesCollection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Variant**|The new data as a string representation of a range contained in the  **[Workbook](Word.ChartData.Workbook.md)** property of the **[ChartData](Word.ChartData.md)** object for the chart.|
| _Rowcol_|Optional| **[XlRowCol](Word.xlrowcol.md)**|One of the enumeration values that specifies whether the new values are in the rows or columns of the specified range.|
| _SeriesLabels_|Optional| **Variant**| **True** if the first row or column contains the name of the data series. **False** if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Word attempts to determine the location of the series name from the contents of the first row or column.|
| _CategoryLabels_|Optional| **Variant**| **True** if the first row or column contains the name of the category labels. **False** if the first row or column contains the first data point of the series. If this argument is omitted, Word attempts to determine the location of the category label from the contents of the first row or column.|
| _Replace_|Optional| **Variant**|If CategoryLabels is  **True** and Replace is **True**, the specified categories replace the categories that currently exist for the series. If Replace is **False**, the existing categories will not be replaced. The default is **False**.|

## Return value

A  **[Series](Word.Series.md)** object that represents the new series.


## Remarks

This method does not actually return a  **Series** object as stated in the Object Browser.


## Example

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


[SeriesCollection Object](Word.SeriesCollection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]