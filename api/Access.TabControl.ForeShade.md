---
title: TabControl.ForeShade property (Access)
keywords: vbaac10.chm14606
f1_keywords:
- vbaac10.chm14606
ms.prod: access
api_name:
- Access.TabControl.ForeShade
ms.assetid: a17b6a7e-d943-f785-459b-43a2c7998ff7
ms.date: 03/01/2019
localization_priority: Normal
---


# TabControl.ForeShade property (Access)

Gets or sets the shade that is applied to the theme color in the **ForeColor** property of the specified object. Read/write **Single**.


## Syntax

_expression_.**ForeShade**

_expression_ A variable that represents a **[TabControl](Access.TabControl.md)** object.


## Remarks

The **ForeShade** property contains a numeric expression that can be used to darken the theme color in the **ForeColor** property. The default value of the **ForeShade** property is 100, which is neutral, and does not change the theme color. 

To darken the color, first determine the percentage by which to darken from 1 to 100, and then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example darkens the **ForeColor** property by 75%.

```vb
Me.ctl.ForeShade=25
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]