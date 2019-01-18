---
title: List.StyleName property (Word)
keywords: vbawd10.chm160563204
f1_keywords:
- vbawd10.chm160563204
ms.prod: word
api_name:
- Word.List.StyleName
ms.assetid: 3d55f975-f6a8-b201-6fd2-e2459fdd048e
ms.date: 06/08/2017
localization_priority: Normal
---


# List.StyleName property (Word)

Returns the name of the style applied to the specified AutoText entry. Read-only  **String**.


## Syntax

 _expression_. `StyleName`

 _expression_ Required. A variable that represents a '[List](Word.List.md)' object.


## Example

This example creates an AutoText entry and then displays the style name of the entry.


```vb
Set myentry = NormalTemplate.AutoTextEntries.Add(Name:="rsvp", _ 
 Range:=Selection.Range) 
MsgBox myentry.StyleName
```


## See also


[List Object](Word.List.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]