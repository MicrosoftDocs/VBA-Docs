---
title: Form.DataEntry property (Access)
keywords: vbaac10.chm13359,vbaac10.chm4316
f1_keywords:
- vbaac10.chm13359,vbaac10.chm4316
ms.prod: access
api_name:
- Access.Form.DataEntry
ms.assetid: 0a970904-10f9-d0c3-24d1-0b988725bb38
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.DataEntry property (Access)

You can use the **DataEntry** property to specify whether a bound form opens to allow data entry only. The **Data Entry** property doesn't determine whether records can be added; it only determines whether existing records are displayed. Read/write **Boolean**.


## Syntax

_expression_.**DataEntry**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

This property can be set in any view.

The **DataEntry** property has an effect only when the **AllowAdditions** property is set to Yes.

Setting the **DataEntry** property to Yes by using Visual Basic has the same effect as choosing **Data Entry** on the **Records** menu. Setting it to No by using Visual Basic is equivalent to choosing **Remove Filter/Sort** on the **Records** menu.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
