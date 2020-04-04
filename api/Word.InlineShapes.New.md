---
title: InlineShapes.New method (Word)
keywords: vbawd10.chm162070728
f1_keywords:
- vbawd10.chm162070728
ms.prod: word
api_name:
- Word.InlineShapes.New
ms.assetid: de83ac06-2b80-69a5-168f-f5f815bfdf11
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShapes.New method (Word)

Inserts an empty, 1-inch-square Word picture object surrounded by a border. This method returns the new graphic as an **InlineShape** object.


## Syntax

_expression_. `New`( `_Range_` )

_expression_ Required. A variable that represents an '[InlineShapes](Word.inlineshapes.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location of the new graphic.|

## Return value

InlineShape


## Example

This example inserts a new, empty picture in the active document and applies a shadow border around the picture.


```vb
Dim ishapeNew As InlineShape 
 
Set ishapeNew = _ 
 ActiveDocument.InlineShapes.New(Range:=Selection.Range) 
 
ishapeNew.Borders.Shadow = True 
ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
```


## See also


[InlineShapes Collection Object](Word.inlineshapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]