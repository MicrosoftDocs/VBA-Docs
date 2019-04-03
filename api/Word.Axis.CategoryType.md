---
title: Axis.CategoryType property (Word)
keywords: vbawd10.chm113049665
f1_keywords:
- vbawd10.chm113049665
ms.prod: word
api_name:
- Word.Axis.CategoryType
ms.assetid: 891a0cce-f5cb-6a8a-6216-fa6aaa1adac9
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.CategoryType property (Word)

Returns or sets the category axis type. Read/write **[XlCategoryType](Word.xlcategorytype.md)**.


## Syntax

_expression_.**CategoryType**

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

You cannot set this property for a value axis.


## Example

The following example sets the category axis for the first chart in the active document to use a time scale, using months as the base unit.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnit = xlMonths 
 End With 
 End If 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]