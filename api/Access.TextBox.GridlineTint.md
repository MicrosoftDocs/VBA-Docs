---
title: TextBox.GridlineTint property (Access)
keywords: vbaac10.chm14636
f1_keywords:
- vbaac10.chm14636
ms.prod: access
api_name:
- Access.TextBox.GridlineTint
ms.assetid: 5dbbd8a7-0942-c39d-b702-a3c0e569e3c1
ms.date: 02/14/2019
localization_priority: Normal
---


# TextBox.GridlineTint property (Access)

Gets or sets the tint applied to the theme color in the **[GridlineColor](access.TextBox.gridlinecolor.md)** property of the specified object. Read/write **Single**.


## Syntax

_expression_.**GridlineTint**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

The **GridlineTint** property contains a numeric expression that can be used to lighten the theme color in the **GridlineColor** property. The default value of the **GridlineTint** property is 100, which is neutral, and does not change the theme color. 

To lighten the color, first determine the percentage by which to lighten from 1 to 100, and then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]