---
title: PivotTable.VisualTotals property (Excel)
keywords: vbaxl10.chm235149
f1_keywords:
- vbaxl10.chm235149
ms.prod: excel
api_name:
- Excel.PivotTable.VisualTotals
ms.assetid: 2bcb64ef-8db8-f62d-5f7d-eb3d5b2fcda5
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.VisualTotals property (Excel)

**True** (default) to enable Online Analytical Processing (OLAP) PivotTables to retotal after an item has been hidden from view. Read/write **Boolean**.


## Syntax

_expression_.**VisualTotals**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

In non-OLAP PivotTables, if you hide an item, the total is recomputed to reflect only the items that remain visible in the PivotTable. In an OLAP PivotTable, the total is computed on the server and is therefore unaffected by whether any items are hidden in the PivotTable view. However, if the **VisualTotals** property is set to **False** for an OLAP PivotTable, the results of the OLAP PivotTable will match those of the non-OLAP PivotTable.

For OLAP PivotTables, a **VisualTotals** property setting of **True** (default) works the same way as described for non-OLAP PivotTables.

The **VisualTotals** property returns **True** for all new PivotTables. However, if you open a workbook in the current version of Microsoft Excel and the PivotTable had been created in a previous version of Excel, the **VisualTotals** property will return **False**.

> [!NOTE] 
> All previously created PivotTables have the **VisualTotals** property set to **False** by default, unless the user changes it, but for all newly created PivotTables, the **VisualTotals** property is set to **True**.


## Example

This example determines if the ability to retotal after an item has been hidden from view is available for OLAP PivotTables and notifies the user. The example assumes that a PivotTable exists on the active worksheet.

```vb
Sub CheckVisualTotals() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine if visual totals is enabled for OLAP PivotTables. 
 If pvtTable.VisualTotals = True Then 
 MsgBox "Ability enabled to re-total after an item " & _ 
 "has been hidden from view." 
 Else 
 MsgBox "Unable to re-total items not hidden from view." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]