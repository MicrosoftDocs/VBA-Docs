---
title: ListEntries.Clear method (Word)
keywords: vbawd10.chm153354342
f1_keywords:
- vbawd10.chm153354342
ms.prod: word
api_name:
- Word.ListEntries.Clear
ms.assetid: 3761ca87-db01-3b84-f1c8-01cc902af5b8
ms.date: 06/08/2017
localization_priority: Normal
---


# ListEntries.Clear method (Word)

Removes all items from a drop-down form field.


## Syntax

_expression_.**Clear**

_expression_ A variable that represents a '[ListEntries](Word.listentries.md)' object.


## Example

This example removes all items from the form field named "Colors" in Sales.doc.


```vb
Documents("Sales.doc").FormFields("Colors") _ 
 .DropDown.ListEntries.Clear
```


## See also


[ListEntries Collection Object](Word.listentries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]