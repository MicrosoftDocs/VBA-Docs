---
title: TextRange2.Paragraphs property (PowerPoint)
ms.assetid: 0f43072e-8f46-4094-b67a-3388b2138c14
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.Paragraphs property (PowerPoint)

Gets a **TextRange2** object that represents the specified subset of text paragraphs. Read-only.


## Syntax

_expression_. `Paragraphs`( `_Start_`, `_Length_` )

 _expression_ An expression that returns a 'TextRange2' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first paragraph in the returned range.|
| _Length_|Optional|**Long**|The number of paragraphs to be returned.|

## Return value

TextRange2


## Remarks

If both  **Start** and **Length** are omitted, the returned range starts with the first paragraph and ends with the last paragraph in the specified range.

If  **Start** is specified but **Length** is omitted, the returned range contains one paragraph.

If  **Length** is specified but **Start** is omitted, the returned range starts with the first paragraph in the specified range.

If  **Start** is greater than the number of paragraphs in the specified text, the returned range starts with the last paragraph in the specified range.

If  **Length** is greater than the number of paragraphs from the specified starting paragraph to the end of the text, the returned range contains all those paragraphs.


## Example

This example formats as italic the first two lines of the second paragraph in shape two on slide one in the active PowerPoint presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _ 
 .TextFrame.TextRange2.Paragraphs(2) _ 
 .Lines(1, 2).Font.Italic = True
```


## See also


[TextRange2 object (PowerPoint)](PowerPoint.textrange2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]