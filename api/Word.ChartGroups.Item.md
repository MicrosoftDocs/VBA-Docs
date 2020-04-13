---
title: ChartGroups.Item method (Word)
keywords: vbawd10.chm77004800
f1_keywords:
- vbawd10.chm77004800
ms.prod: word
api_name:
- Word.ChartGroups.Item
ms.assetid: 0d78e50d-f2e1-1617-a563-65cc48ca2c30
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroups.Item method (Word)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a '[ChartGroups](Word.ChartGroups.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number for the object.|

## Return value

A **[ChartGroup](Word.ChartGroup.md)** object contained by the collection.


## Example

The following example adds drop lines to chart group one for the first chart group of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups.Item(1).HasDropLines = True 
 End If 
End With
```


## See also


[ChartGroups Object](Word.ChartGroups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]