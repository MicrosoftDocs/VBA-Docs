---
title: AutoTextEntry.StyleName property (Word)
keywords: vbawd10.chm154533891
f1_keywords:
- vbawd10.chm154533891
ms.prod: word
api_name:
- Word.AutoTextEntry.StyleName
ms.assetid: 0bcb48b2-c131-4bff-732e-ec221f24e463
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoTextEntry.StyleName property (Word)

Returns the name of the style applied to the specified AutoText entry. Read-only  **String**.


## Syntax

_expression_. `StyleName`

_expression_ A variable that represents a '[AutoTextEntry](Word.AutoTextEntry.md)' object.


## Example

This example creates an AutoText entry and then displays the style name of the entry.


```vb
Set myentry = NormalTemplate.AutoTextEntries.Add(Name:="rsvp", _ 
 Range:=Selection.Range) 
MsgBox myentry.StyleName
```


## See also


[AutoTextEntry Object](Word.AutoTextEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]