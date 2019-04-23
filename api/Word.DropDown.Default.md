---
title: DropDown.Default property (Word)
keywords: vbawd10.chm153419777
f1_keywords:
- vbawd10.chm153419777
ms.prod: word
api_name:
- Word.DropDown.Default
ms.assetid: aaf2920c-1077-3c5d-f80e-4ad119b3ae2f
ms.date: 06/08/2017
localization_priority: Normal
---


# DropDown.Default property (Word)

Returns or sets a  **Long** that represents the default drop-down item. Read/write.


## Syntax

_expression_. `Default`

_expression_ Required. A variable that represents a '[DropDown](Word.DropDown.md)' object.


## Remarks

The first item in a drop-down form field is 1, the second item is 2, and so on.


## Example

This example sets the default item for the drop-down form field named "Colors" in Sales.doc.


```vb
Documents("Sales.doc").FormFields("Colors").DropDown _ 
 .Default = 2
```


## See also


[DropDown Object](Word.DropDown.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]