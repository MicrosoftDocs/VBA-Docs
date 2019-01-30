---
title: TextRange2.Characters property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Characters
ms.assetid: 9b264529-e538-4480-e629-822d5056f148
ms.date: 01/25/2019
localization_priority: Normal
---


# TextRange2.Characters property (Office)

Read-only.


## Syntax

_expression_.**Characters** (_Start_, _Length_)

_expression_ An expression that returns a **[TextRange2](Office.TextRange2.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first character in the returned range.|
| _Length_|Optional|**Long**|The number of characters to be returned.|

## Return value

TextRange2


## Remarks

If both _Start_ and _Length_ are omitted, the returned range starts with the first character and ends with the last paragraph in the specified range.

If _Start_ is specified but _Length_ is omitted, the returned range contains one character.

If _Length_ is specified but _Start_ is omitted, the returned range starts with the first character in the specified range.

If _Start_ is greater than the number of characters in the specified text, the returned range starts with the last character in the specified range.

If _Length_ is greater than the number of characters from the specified starting character to the end of the text, the returned range contains all those characters.


## Example

This example sets the text for shape two on slide one in the active presentation, and then makes the second character a subscript character with a 20-percent offset.


```vb
Dim charRange As TextRange2 
With Application.ActivePresentation.Slides(1).Shapes(2) 
 Set charRange = .TextFrame.TextRange2.InsertBefore("H2O") 
 charRange.Characters(2).Font.BaselineOffset = -0.2 
End With 

```


## See also

- [TextRange2 object members](overview/Library-Reference/textrange2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]