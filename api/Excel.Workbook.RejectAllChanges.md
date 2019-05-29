---
title: Workbook.RejectAllChanges method (Excel)
keywords: vbaxl10.chm199178
f1_keywords:
- vbaxl10.chm199178
ms.prod: excel
api_name:
- Excel.Workbook.RejectAllChanges
ms.assetid: a53670da-370c-9ac8-75b8-008625495c2b
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.RejectAllChanges method (Excel)

Rejects all changes in the specified shared workbook.


## Syntax

_expression_.**RejectAllChanges** (_When_, _Who_, _Where_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _When_|Optional| **Variant**|Specifies when all the changes are rejected.|
| _Who_|Optional| **Variant**|Specifies by whom all the changes are rejected.|
| _Where_|Optional| **Variant**|Specifies where all the changes are rejected.|

## Example

This example rejects all changes in the active workbook.

```vb
ActiveWorkbook.RejectAllChanges
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]