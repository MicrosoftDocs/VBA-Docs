---
title: ObjectFrame.BorderShade Property (Access)
keywords: vbaac10.chm14603
f1_keywords:
- vbaac10.chm14603
ms.prod: access
api_name:
- Access.ObjectFrame.BorderShade
ms.assetid: de92000e-95bb-12df-68ef-5ada76553e97
ms.date: 06/08/2017
---


# ObjectFrame.BorderShade Property (Access)

Gets or sets the shade that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **BorderShade**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

The  **BorderShade** property contains a numeric expression that can be used to darken the theme color in the **BorderColor** property. The default value of the **BorderShade** property is 100, which is neutral, and does not change the theme color. To darken the color, first determine the percentage by which to darken from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## Example

The following code example darkens  **BorderColor** by 75%.


```vb
Me.ctl.BorderShade=25
```


## See also


#### Concepts


[ObjectFrame Object](Access.ObjectFrame.md)

