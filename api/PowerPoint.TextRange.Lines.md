---
title: TextRange.Lines method (PowerPoint)
keywords: vbapp10.chm569014
f1_keywords:
- vbapp10.chm569014
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Lines
ms.assetid: 8e9f344b-2e74-5a9d-06e8-3e6ff9ca6bd0
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.Lines method (PowerPoint)

Returns a **[TextRange](PowerPoint.TextRange.md)** object that represents the specified subset of text lines. For information about counting or looping through the lines in a text range, see the **[TextRange](PowerPoint.TextRange.md)** object.


## Syntax

_expression_. `Lines`( `_Start_`, `_Length_` )

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first line in the returned range.|
| _Length_|Optional|**Long**|The number of lines to be returned.|

## Return value

TextRange


## Remarks

If both Start and Length are omitted, the returned range starts with the first line and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one line.

If Length is specified but Start is omitted, the returned range starts with the first line in the specified range.

If Start is greater than the number of lines in the specified text, the returned range starts with the last line in the specified range.

If Length is greater than the number of lines from the specified starting line to the end of the text, the returned range contains all those lines.


## Example

This example formats as italic the first two lines of the second paragraph in shape two on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _
    .TextFrame.TextRange.Paragraphs(2) _
    .Lines(1, 2).Font.Italic = True
```


## See also


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]