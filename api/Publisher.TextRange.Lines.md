---
title: TextRange.Lines method (Publisher)
keywords: vbapb10.chm5308455
f1_keywords:
- vbapb10.chm5308455
ms.prod: publisher
api_name:
- Publisher.TextRange.Lines
ms.assetid: 56862090-b2ff-403b-d016-e37108d5ccc1
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Lines method (Publisher)

Returns a **TextRange** object that represents the specified lines.


## Syntax

_expression_.**Lines** (_Start_, _Length_)

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Start_|Required| **Long**|The first line in the returned range.|
|_Length_|Optional| **Long**|The number of lines to be returned. The default is 1.|

## Return value

TextRange


## Remarks

If _Start_ is greater than the number of lines in the specified text, the returned range starts with the last line in the specified range.

If _Length_ is greater than the number of lines from the specified starting line to the end of the text, the returned range contains all those lines.


## Example

This example replaces the first three lines of the first shape on the first page with the specified string.

```vb
Sub ReplaceLines() 
 Dim rngText As TextRange 
 Set rngText = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Lines(Start:=1, Length:=3) 
 
 rngText.Text = "This is replacement text." & vbCrLf 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]