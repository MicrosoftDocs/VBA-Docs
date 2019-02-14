---
title: ObjectFrame.GridlineShade property (Access)
keywords: vbaac10.chm14637
f1_keywords:
- vbaac10.chm14637
ms.prod: access
api_name:
- Access.ObjectFrame.GridlineShade
ms.assetid: 837f3c0b-5597-7abd-e580-c92f099d4448
ms.date: 02/14/2019
localization_priority: Normal
---


# ObjectFrame.GridlineShade property (Access)

Gets or sets the shade applied to the theme color in the **[GridlineColor](access.objectframe.gridlinecolor.md)** property of the specified object. Read/write **Single**.


## Syntax

_expression_.**GridlineShade**

_expression_ A variable that represents an **[ObjectFrame](Access.ObjectFrame.md)** object.


## Remarks

The **GridlineShade** property contains a numeric expression that can be used to darken the theme color in the **GridlineColor** property. The default value of the **GridlineShade** property is 100, which is neutral, and does not change the theme color. 

To darken the color, first determine the percentage by which to darken from 1 to 100, and then subtract that value as a whole number from 100 and store the remainder. For example, to darken the theme color shade by 75%, subtract 75 from 100 and store the remainder, which is 25.

This property is not surfaced in the property sheet.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]