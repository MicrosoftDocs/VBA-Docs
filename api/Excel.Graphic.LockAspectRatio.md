---
title: Graphic.LockAspectRatio property (Excel)
keywords: vbaxl10.chm694082
f1_keywords:
- vbaxl10.chm694082
ms.prod: excel
api_name:
- Excel.Graphic.LockAspectRatio
ms.assetid: d38851e4-7ca6-bb1f-4b16-03fe78fae726
ms.date: 04/26/2019
localization_priority: Normal
---


# Graphic.LockAspectRatio property (Excel)

**True** if the specified shape retains its original proportions when you resize it. **False** if you can change the height and width of the shape independently of one another when you resize it. Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**LockAspectRatio**

_expression_ A variable that represents a **[Graphic](Excel.Graphic.md)** object.


## Remarks

With **msoFalse**, you can change the height and width of the shape independently of one another when you resize it.

With **msoTrue**, the specified shape retains its original proportions when you resize it.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]