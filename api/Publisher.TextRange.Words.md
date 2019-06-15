---
title: TextRange.Words method (Publisher)
keywords: vbapb10.chm5308456
f1_keywords:
- vbapb10.chm5308456
ms.prod: publisher
api_name:
- Publisher.TextRange.Words
ms.assetid: df812db2-98ca-848b-7922-6905cb71124c
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Words method (Publisher)

Returns a **TextRange** object that represents the specified subset of text words.


## Syntax

_expression_.**Words** (_Start_, _Length_)

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Start_|Required| **Long**|The first word in the returned range.|
|_Length_|Optional| **Long**|The number of words to be returned. The default is 1.|

## Return value

TextRange


## Remarks

If _Length_ is omitted, the returned range contains one word.

If _Start_ is greater than the number of words in the specified text, the returned range starts with the last word in the specified range.

If _Length_ is greater than the number of words from the specified starting word to the end of the text, the returned range contains all those words.


## Example

This example formats as bold the second, third, and fourth words in shape two on page one of the active publication.

```vb
Application.ActiveDocument.Pages(1).Shapes(2) _ 
 .TextFrame.TextRange.Words(Start:=2, Length:=3) _ 
 .Font.Bold = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]