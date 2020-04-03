---
title: DataLabels.Item method (Word)
keywords: vbawd10.chm207486976
f1_keywords:
- vbawd10.chm207486976
ms.prod: word
api_name:
- Word.DataLabels.Item
ms.assetid: 792b63a5-e4e9-c026-e94d-0f0349d113dc
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabels.Item method (Word)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[DataLabels](Word.DataLabels.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number for the object.|

## Return value

A **[DataLabel](Word.DataLabel.md)** object contained by the collection.


## Example

The following example sets the number format for the fifth data label in the first series for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels.Item(5). _ 
 NumberFormat = "0.000" 
 End If 
End With 

```


## See also


[DataLabels Object](Word.DataLabels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]