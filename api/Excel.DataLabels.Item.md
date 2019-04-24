---
title: DataLabels.Item method (Excel)
keywords: vbaxl10.chm584106
f1_keywords:
- vbaxl10.chm584106
ms.prod: excel
api_name:
- Excel.DataLabels.Item
ms.assetid: bc45ebcc-00f0-c253-0d68-002d8f20d750
ms.date: 04/23/2019
localization_priority: Normal
---


# DataLabels.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[DataLabels](Excel.DataLabels(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number for the object.|

## Return value

A **[DataLabel](Excel.DataLabel(object).md)** object contained by the collection.


## Example

This example sets the number format for the fifth data label in series one in embedded chart one on worksheet one.

```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).DataLabels.Item(5).NumberFormat = "0.000"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]