---
title: TextRange2.Paragraphs property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Paragraphs
ms.assetid: 15479f9e-f261-7ea6-0460-861ccea08440
ms.date: 01/25/2019
localization_priority: Normal
---


# TextRange2.Paragraphs property (Office)

Gets a **TextRange2** object that represents the specified subset of text paragraphs. Read-only.


## Syntax

_expression_.**Paragraphs** (_Start_, _Length_)

_expression_ An expression that returns a **[TextRange2](Office.TextRange2.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first paragraph in the returned range.|
| _Length_|Optional|**Long**|The number of paragraphs to be returned.|

## Return value

TextRange2


## Remarks

If both _Start_ and _Length_ are omitted, the returned range starts with the first paragraph and ends with the last paragraph in the specified range.

If _Start_ is specified but _Length_ is omitted, the returned range contains one paragraph.

If _Length_ is specified but _Start_ is omitted, the returned range starts with the first paragraph in the specified range.

If _Start_ is greater than the number of paragraphs in the specified text, the returned range starts with the last paragraph in the specified range.

If _Length_ is greater than the number of paragraphs from the specified starting paragraph to the end of the text, the returned range contains all those paragraphs.


## Example

This example formats as italic the first two lines of the second paragraph in shape two on slide one in the active PowerPoint presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _ 
 .TextFrame.TextRange2.Paragraphs(2) _ 
 .Lines(1, 2).Font.Italic = True
```


## See also

- [TextRange2 object members](overview/Library-Reference/textrange2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]