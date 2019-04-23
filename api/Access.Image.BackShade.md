---
title: Image.BackShade property (Access)
keywords: vbaac10.chm14633
f1_keywords:
- vbaac10.chm14633
ms.prod: access
api_name:
- Access.Image.BackShade
ms.assetid: 899c5320-a2ef-7861-2905-fc08f5b7a1fb
ms.date: 02/28/2019
localization_priority: Normal
---


# Image.BackShade property (Access)

Gets or sets the shade applied to the theme color in the **BackColor** property of the specified object. Read/write **Single**.


## Syntax

_expression_.**BackShade**

_expression_ A variable that represents an **[Image](Access.Image.md)** object.


## Remarks

The **BackShade** property contains a numeric expression that can be used to darken the theme color in the **BackColor** property. The default value of the **BackShade** property is 100, which is neutral, and does not change the theme color. 

To darken the color, first determine the percentage by which to darken from 1 to 100, and then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example darkens the **BackColor** property by 75%.

```vb
Me.ctl.BackShade=25
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]