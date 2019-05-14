---
title: SeriesCollection.Extend method (Excel)
keywords: vbaxl10.chm580076
f1_keywords:
- vbaxl10.chm580076
ms.prod: excel
api_name:
- Excel.SeriesCollection.Extend
ms.assetid: 85f2658f-b7b3-e086-da27-5127f1ea4ff7
ms.date: 05/14/2019
localization_priority: Normal
---


# SeriesCollection.Extend method (Excel)

Adds new data points to an existing series collection.


## Syntax

_expression_.**Extend** (_Source_, _RowCol_, _CategoryLabels_)

_expression_ A variable that represents a **[SeriesCollection](Excel.SeriesCollection.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Variant**|The new data to be added to the **SeriesCollection** object as a **[Range](Excel.Range(object).md)** object.|
| _RowCol_|Optional| **Variant**|Specifies whether the new values are in the rows or columns of the given range source. Can be one of the following **[XlRowCol](Excel.XlRowCol.md)** constants: **xlRows** or **xlColumns**. If this argument is omitted, Microsoft Excel attempts to determine where the values are by the size and orientation of the selected range or by the dimensions of the array.|
| _CategoryLabels_|Optional| **Variant**| **True** to have the first row or column contain the name of the category labels. **False** to have the first row or column contain the first data point of the series. If this argument is omitted, Excel attempts to determine the location of the category label from the contents of the first row or column.|

## Return value

Variant


## Remarks

This method is not available for PivotChart reports.


## Example

This example extends the series on Chart1 by adding the data in cells B1:B6 on Sheet1.

```vb
Charts("Chart1").SeriesCollection.Extend _ 
        Source:=Worksheets("Sheet1").Range("B1:B6") 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]