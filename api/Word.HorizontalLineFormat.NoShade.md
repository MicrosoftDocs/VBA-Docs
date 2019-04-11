---
title: HorizontalLineFormat.NoShade property (Word)
keywords: vbawd10.chm165543939
f1_keywords:
- vbawd10.chm165543939
ms.prod: word
api_name:
- Word.HorizontalLineFormat.NoShade
ms.assetid: 90728761-cdfa-fd2c-db00-44ca78a34017
ms.date: 06/08/2017
localization_priority: Normal
---


# HorizontalLineFormat.NoShade property (Word)

 **True** if Microsoft Word draws the specified horizontal line without 3D shading. Read/write **Boolean**.


## Syntax

_expression_. `NoShade`

 _expression_ An expression that returns a '[HorizontalLineFormat](Word.HorizontalLineFormat.md)' object.


## Remarks

You can only use this property with horizontal lines that are not based on an existing image file.


## Example

This example adds a horizontal line without any 3D shading.


```vb
Selection.InlineShapes.AddHorizontalLineStandard 
ActiveDocument.InlineShapes(1) _ 
 .HorizontalLineFormat.NoShade = True
```


## See also


[HorizontalLineFormat Object](Word.HorizontalLineFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]