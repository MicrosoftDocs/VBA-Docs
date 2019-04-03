---
title: TextRange2.Lines property (PowerPoint)
ms.assetid: 09b52bda-e1ab-4cf2-bf38-25a156bf4c2e
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.Lines property (PowerPoint)

Returns a TextRange2 object that represents the specified subset of text lines. Read-only.


## Syntax

_expression_. `Lines`( `_Start_`, `_Length_` )

 _expression_ An expression that returns a 'TextRange2' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first line in the returned range.|
| _Length_|Optional|**Long**|The number of lines to be returned.|

## Return value

TextRange2


## Remarks

If both Start and Length are omitted, the returned range starts with the first line and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one line.

If Length is specified but Start is omitted, the returned range starts with the first line in the specified range.

If Start is greater than the number of lines in the specified text, the returned range starts with the last line in the specified range.

If Length is greater than the number of lines from the specified starting line to the end of the text, the returned range contains all those lines.


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