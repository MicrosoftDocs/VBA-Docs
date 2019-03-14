---
title: Form.SelTop property (Access)
keywords: vbaac10.chm13470
f1_keywords:
- vbaac10.chm13470
ms.prod: access
api_name:
- Access.Form.SelTop
ms.assetid: 5503187c-09ea-222e-5db2-f3c2298f34dc
ms.date: 03/15/2019
localization_priority: Normal
---


# Form.SelTop property (Access)

You can use the **SelTop** property to specify or determine which row (record) is topmost in the current selection rectangle in a table, query, or form datasheet, or which selected record is topmost in a continuous form. Read/write **Long**.


## Syntax

_expression_.**SelTop**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **SelTop** property returns a value between 1 and the number of records in the datasheet or continuous form.

If there's no selection, the value returned by this property is the row and column of the cell with the focus.

If you've selected one or more columns (using the column headings), you can't change the setting of the **SelTop** property.

You can use these properties with the **SelHeight** and **SelWidth** properties to specify or determine the actual size of the selection rectangle. 

The **SelTop** and **SelLeft** properties determine the position of the upper-left corner of the selection rectangle. 

The **SelHeight** and **SelWidth** properties determine the lower-right corner of the selection rectangle.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]