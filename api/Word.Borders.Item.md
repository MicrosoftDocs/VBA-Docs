---
title: Borders.Item method (Word)
keywords: vbawd10.chm154927104
f1_keywords:
- vbawd10.chm154927104
ms.prod: word
api_name:
- Word.Borders.Item
ms.assetid: ac2b9108-5ae1-e875-f6a0-47a8c2175fe1
ms.date: 06/08/2017
localization_priority: Normal
---


# Borders.Item method (Word)

Returns a border in a range or selection.


## Syntax

 _expression_. `Item`( `_Index_` )

 _expression_ Required. A variable that represents a '[Borders](Word.borders.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **WdBorderType**|The border to be returned.|

## Return value

Border


## Example

This example inserts a double border above the first paragraph in the active document.


```vb
Sub BorderItem() 
 ActiveDocument.Paragraphs(1).Borders.Item(wdBorderTop) _ 
 .LineStyle = wdLineStyleDouble 
End Sub
```


## See also


[Borders Collection Object](Word.borders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]