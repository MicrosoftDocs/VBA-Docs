---
title: Filters.Item property (Excel)
keywords: vbaxl10.chm540075
f1_keywords:
- vbaxl10.chm540075
ms.prod: excel
api_name:
- Excel.Filters.Item
ms.assetid: a24c9aeb-b253-c11a-29dc-c4a2bba86e21
ms.date: 04/26/2019
localization_priority: Normal
---


# Filters.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Filters](Excel.Filters.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object.|


## Example

The following example sets a variable to the value of the **On** property of the filter for the first column in the filtered range on the Crew worksheet.

```vb
Set w = Worksheets("Crew") 
If w.AutoFilterMode Then 
 filterIsOn = w.AutoFilter.Filters.Item(1).On 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]