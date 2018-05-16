---
title: EmptyCell.Height Property (Access)
keywords: vbaac10.chm14310
f1_keywords:
- vbaac10.chm14310
ms.prod: access
api_name:
- Access.EmptyCell.Height
ms.assetid: 35a90799-a5fb-62da-1e75-f4f2396cac2e
ms.date: 06/08/2017
---


# EmptyCell.Height Property (Access)

Gets or sets the height of the specified object in twips. Read/write  **Integer**.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents an **EmptyCell** object.


## Remarks

For report controls, you can set the  **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

Microsoft Access automatically sets the  **Height** property when you create or size a control or when you size a window in form Design View or report Design view.

The height of controls is measured from the center of their borders so controls with different border widths align correctly. 


## See also


#### Concepts


[EmptyCell Object](Access.EmptyCell.md)

