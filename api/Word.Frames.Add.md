---
title: Frames.Add method (Word)
keywords: vbawd10.chm153813092
f1_keywords:
- vbawd10.chm153813092
ms.prod: word
api_name:
- Word.Frames.Add
ms.assetid: e9b25f79-b95d-fcd4-f88c-a32b5f83f3dc
ms.date: 06/08/2017
localization_priority: Normal
---


# Frames.Add method (Word)

Returns a Frame object that represents a new frame added to a range, selection, or document.


## Syntax

_expression_.**Add** (_Range_)

 _expression_ An expression that returns a '[Frames](Word.Frames.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **[RANGE]**|The range that you want the frame to surround.|

## Return value

Frame


## Example

This example adds a frame around the selection.


```vb
ActiveDocument.Frames.Add Range:=Selection.Range
```

This example adds a frame around the third paragraph in the selection.




```vb
Set myFrame = Selection.Frames _ 
 .Add(Range:=Selection.Paragraphs(3).Range)
```


## See also


[Frames Object](Word.Frames.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]