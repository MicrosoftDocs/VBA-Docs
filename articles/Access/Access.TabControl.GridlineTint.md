---
title: TabControl.GridlineTint Property (Access)
keywords: vbaac10.chm14636
f1_keywords:
- vbaac10.chm14636
ms.prod: access
api_name:
- Access.TabControl.GridlineTint
ms.assetid: 38c2d0ea-6832-3ea3-c0fd-74d2135cfafd
ms.date: 06/08/2017
---


# TabControl.GridlineTint Property (Access)

Gets or sets the tint applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.


## Syntax

 _expression_. **GridlineTint**

 _expression_ A variable that represents a **TabControl** object.


## Remarks

The  **GridlineTint** property contains a numeric expression that can be used to lighten the theme color in the **GridlineColor** property. The default value of the **GridlineTint** property is 100, which is neutral, and does not change the theme color. To lighten the color, first determine the percentage by which to lighten from 1 to 100, then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[TabControl Object](Access.TabControl.md)

