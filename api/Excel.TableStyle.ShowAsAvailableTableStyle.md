---
title: TableStyle.ShowAsAvailableTableStyle property (Excel)
keywords: vbaxl10.chm839078
f1_keywords:
- vbaxl10.chm839078
ms.prod: excel
api_name:
- Excel.TableStyle.ShowAsAvailableTableStyle
ms.assetid: cf5c7b9c-6ed9-e26e-4b31-614ede2a4a12
ms.date: 05/17/2019
localization_priority: Normal
---


# TableStyle.ShowAsAvailableTableStyle property (Excel)

Returns or sets a table style shown as available in the table styles gallery. Read/write **Boolean**.


## Syntax

_expression_.**ShowAsAvailableTableStyle**

_expression_ A variable that represents a **[TableStyle](Excel.TableStyle.md)** object.


## Remarks

If **True**, this style is shown in the gallery for table styles.

You can set this property to **False** even when the style is already applied to a table. In this case, the gallery will not show the style, and when the active cell is in that table, no style is shown as selected.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]