---
title: Workbook.UpdateRemoteReferences property (Excel)
keywords: vbaxl10.chm199161
f1_keywords:
- vbaxl10.chm199161
ms.prod: excel
api_name:
- Excel.Workbook.UpdateRemoteReferences
ms.assetid: 055c1a88-c189-ddd3-c9b2-9458817cec90
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.UpdateRemoteReferences property (Excel)

**True** if Microsoft Excel updates remote references in the workbook. Read/write **Boolean**.


## Syntax

_expression_.**UpdateRemoteReferences**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

When a new workbook is created, the default value for the **UpdateRemoteReferences** property is **True** and dynamic data exchange (DDE) links and OLE links update automatically. If the value is **False**, DDE links and OLE links do not update automatically or during recalculation.


## Example

This example causes remote references to update automatically in the active workbook.

```vb
ActiveWorkbook.UpdateRemoteReferences = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]