---
title: ObjectFrame.BackShade property (Access)
keywords: vbaac10.chm14633
f1_keywords:
- vbaac10.chm14633
ms.prod: access
api_name:
- Access.ObjectFrame.BackShade
ms.assetid: 68800e85-9dfa-958d-e87d-1241be551f90
ms.date: 02/28/2019
localization_priority: Normal
---


# ObjectFrame.BackShade property (Access)

Gets or sets the shade applied to the theme color in the **BackColor** property of the specified object. Read/write **Single**.


## Syntax

_expression_.**BackShade**

_expression_ A variable that represents an **[ObjectFrame](Access.ObjectFrame.md)** object.


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