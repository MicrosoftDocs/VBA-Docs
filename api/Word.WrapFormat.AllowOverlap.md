---
title: WrapFormat.AllowOverlap property (Word)
keywords: vbawd10.chm163774570
f1_keywords:
- vbawd10.chm163774570
ms.prod: word
api_name:
- Word.WrapFormat.AllowOverlap
ms.assetid: b224d70d-0128-cfec-39f2-97fd12b0c5ca
ms.date: 06/08/2017
localization_priority: Normal
---


# WrapFormat.AllowOverlap property (Word)

Returns or sets a value that specifies whether a given shape can overlap other shapes. Read/write **Long**.


## Syntax

_expression_.**AllowOverlap**

_expression_ A variable that represents a **[WrapFormat](Word.WrapFormat.md)** object.


## Remarks

This property can be set to either **True** or **False**. Because HTML doesn't support overlapping tables or shapes, **AllowOverlap** is ignored in web layout view.


## Example

This example specifies that the first shape in the active document can overlap other shapes.

```vb
ActiveDocument.Shapes(1).WrapFormat.AllowOverlap = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]