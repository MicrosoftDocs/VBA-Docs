---
title: FormatConditions.Item method (Excel)
keywords: vbaxl10.chm510074
f1_keywords:
- vbaxl10.chm510074
ms.prod: excel
api_name:
- Excel.FormatConditions.Item
ms.assetid: 62b8bef8-94ae-5cfa-0af7-bd6a311f9cb2
ms.date: 04/26/2019
localization_priority: Normal
---


# FormatConditions.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[FormatConditions](Excel.FormatConditions.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

An Object value that represents an object contained by the collection.


## Example

This example sets format properties for an existing conditional format for cells E1:E10.

```vb
With Worksheets(1).Range("e1:e10").FormatConditions.Item(1) 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]