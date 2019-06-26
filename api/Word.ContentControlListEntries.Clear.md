---
title: ContentControlListEntries.Clear method (Word)
keywords: vbawd10.chm230948968
f1_keywords:
- vbawd10.chm230948968
ms.prod: word
api_name:
- Word.ContentControlListEntries.Clear
ms.assetid: baaae83d-98ad-18ee-9302-632fbf5271fe
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControlListEntries.Clear method (Word)

Clears all items from a drop-down list or combo box content control.


## Syntax

_expression_.**Clear**

 _expression_ An expression that returns a [ContentControlListEntries](./Word.ContentControlListEntries.md) object.


## Example

The following code example clears all items from the first content control in the active document.


> [!NOTE] 
> The following code example assumes that the first content control in the active document is a drop-down list or combo box.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls(1) 
 
objCC.DropdownListEntries.Clear
```


## See also


[ContentControlListEntries Collection](Word.ContentControlListEntries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]