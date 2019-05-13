---
title: SeriesCollection.Paste method (Excel)
keywords: vbaxl10.chm580079
f1_keywords:
- vbaxl10.chm580079
ms.prod: excel
api_name:
- Excel.SeriesCollection.Paste
ms.assetid: 460644ba-e682-d4dd-4832-f9f18fb6389b
ms.date: 05/14/2019
localization_priority: Normal
---


# SeriesCollection.Paste method (Excel)

Pastes data from the Clipboard into the specified series collection.


## Syntax

_expression_.**Paste** (_RowCol_, _SeriesLabels_, _CategoryLabels_, _Replace_, _NewSeries_)

_expression_ A variable that represents a **[SeriesCollection](Excel.SeriesCollection.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RowCol_|Optional| **[XlRowCol](Excel.XlRowCol.md)**| Specifies whether the values corresponding to a particular data series are in rows or columns.|
| _SeriesLabels_|Optional| **Variant**| **True** to use the contents of the cell in the first column of each row (or the first row of each column) as the name of the data series in that row (or column). **False** to use the contents of the cell in the first column of each row (or the first row of each column) as the first data point in the data series. The default value is **False**.|
| _CategoryLabels_|Optional| **Variant**| **True** to use the contents of the first row (or column) of the selection as the categories for the chart. **False** to use the contents of the first row (or column) as the first data series in the chart. The default value is **False**.|
| _Replace_|Optional| **Variant**| **True** to apply categories while replacing existing categories with information from the copied range. **False** to insert new categories without replacing any old ones. The default value is **True**.|
| _NewSeries_|Optional| **Variant**| **True** to paste the data as a new series. **False** to paste the data as new points in an existing series. The default value is **True**.|

## Return value

Variant


## Example

This example pastes data from the Clipboard into a new series on Chart1.

```vb
Worksheets("Sheet1").Range("C1:C5").Copy 
Charts("Chart1").SeriesCollection.Paste
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]