---
title: TextRange.Paragraphs method (Publisher)
keywords: vbapb10.chm5308454
f1_keywords:
- vbapb10.chm5308454
ms.prod: publisher
api_name:
- Publisher.TextRange.Paragraphs
ms.assetid: 895c32cf-cdbe-74b0-ab47-6ae63d1bdea0
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Paragraphs method (Publisher)

Returns a **TextRange** object that represents the specified paragraphs.


## Syntax

_expression_.**Paragraphs** (_Start_, _Length_)

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Start_|Required| **Long**|The first paragraph in the returned range.|
|_Length_|Optional| **Long**|The number of paragraphs to be returned. The default is 1.|

## Return value

TextRange


## Remarks

If _Length_ is omitted, the returned range contains one paragraph.

If _Length_ is greater than the number of paragraphs from the specified starting paragraph to the end of the text, the returned range contains all those paragraphs.


## Example

This example formats as indents the first line of the selected paragraph.

```vb
Sub FormatCurrentParagraph() 
 Selection.TextRange.Paragraphs(Start:=1).ParagraphFormat _ 
 .FirstLineIndent = InchesToPoints(0.5) 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]