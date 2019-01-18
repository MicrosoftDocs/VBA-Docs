---
title: TextRange2.Lines property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Lines
ms.assetid: 5e20f089-c345-e22a-c136-483d13f7f658
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange2.Lines property (Office)

Returns a TextRange2 object that represents the specified subset of text lines. Read-only.


## Syntax

_expression_. `Lines`( `_Start_`, `_Length_` )

 _expression_ An expression that returns a [TextRange2](Office.TextRange2.md) object.


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


[TextRange2 Object](Office.TextRange2.md)



[TextRange2 Object Members](./overview/Library-Reference/textrange2-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]