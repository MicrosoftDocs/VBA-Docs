---
title: OptionGroup.BackTint Property (Access)
keywords: vbaac10.chm14632
f1_keywords:
- vbaac10.chm14632
ms.prod: access
api_name:
- Access.OptionGroup.BackTint
ms.assetid: 4e33a712-af8f-bffa-f6c8-0502fb292813
ms.date: 06/08/2017
---


# OptionGroup.BackTint Property (Access)

Gets or sets the tint that is applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **BackTint**

 _expression_ A variable that represents an **OptionGroup** object.


## Remarks

The  **BackTint** property contains a numeric expression that can be used to lighten the theme color in the **BackColor** property. The default value of the **BackTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example lightens  **BackColor** by 75%.


```vb
Me.ctl.BackTint=25
```


## See also


#### Concepts


[OptionGroup Object](Access.OptionGroup.md)

