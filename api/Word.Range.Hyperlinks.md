---
title: Range.Hyperlinks property (Word)
keywords: vbawd10.chm157155484
f1_keywords:
- vbawd10.chm157155484
ms.prod: word
api_name:
- Word.Range.Hyperlinks
ms.assetid: c8eb84af-b090-82ee-8001-b251c6cc1f24
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Hyperlinks property (Word)

Returns a  **Hyperlinks** collection that represents all the hyperlinks in the specified range. Read-only.


## Syntax

_expression_.**Hyperlinks**

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the name of every hyperlink in the first ten paragraphs in the active document.


```vb
Dim objLink As Hyperlink 
Dim objRange As Range 
 
Set objRange = ActiveDocument.Range( _ 
 Paragraphs(1).Range.Start, _ 
 Paragraphs(10).Range.End) 
 
For Each objLink In objRange.Hyperlinks 
 If InStr(LCase(objLink.Address), "microsoft") <> 0 Then 
 MsgBox objLink.Name 
 End If 
Next objLink
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]