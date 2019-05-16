---
title: TableStyle.ShowAsAvailablePivotTableStyle property (Excel)
keywords: vbaxl10.chm839079
f1_keywords:
- vbaxl10.chm839079
ms.prod: excel
api_name:
- Excel.TableStyle.ShowAsAvailablePivotTableStyle
ms.assetid: c9439773-e9e2-d642-ed80-4b44b7e79130
ms.date: 05/17/2019
localization_priority: Normal
---


# TableStyle.ShowAsAvailablePivotTableStyle property (Excel)

Sets or returns whether a style is shown in the gallery for PivotTable styles. Read/write **Boolean**.


## Syntax

_expression_.**ShowAsAvailablePivotTableStyle**

_expression_ A variable that represents a **[TableStyle](Excel.TableStyle.md)** object.


## Remarks

The property returns **True** if the style is shown in the gallery for PivotTable styles.

> [!NOTE] 
> Users can set the **ShowAsAvailableTableStyle** or **ShowAsAvailablePivotTableStyle** properties to **False** even when the style is already applied to a table or PivotTable. In this case, the gallery will not show the style, and no style is shown as selected when the active cell is in the table or the PivotTable.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]