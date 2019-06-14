---
title: TextRange.Characters method (Publisher)
keywords: vbapb10.chm5308425
f1_keywords:
- vbapb10.chm5308425
ms.prod: publisher
api_name:
- Publisher.TextRange.Characters
ms.assetid: e851767e-12b2-ad77-071b-9d27bbf0d637
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Characters method (Publisher)

Returns a **TextRange** object that represents the specified subset of text characters.


## Syntax

_expression_.**Characters** (_Start_, _Length_)

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Start_|Required| **Long**|The first character in the returned range.|
|_Length_|Optional| **Long**|The number of characters to be returned. Default is 1.|

## Return value

TextRange


## Remarks

If _Start_ is greater than the number of characters in the specified text, the returned range starts with the last character in the specified range.

If _Length_ is greater than the number of characters from the specified starting character to the end of the text, the returned range contains all those characters.


## Example

This example sets the text for the first shape on the first page in the active document, and then sets the font of the first two characters to 15 points and bold.

```vb
Sub CharRange() 
 Dim rngCharacters As TextRange 
 Set rngCharacters = Application.ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.InsertBefore(NewText:="Hello World.") 
 With rngCharacters.Characters(Start:=1, Length:=2).Font 
 .Size = 15 
 .Bold = msoTrue 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]