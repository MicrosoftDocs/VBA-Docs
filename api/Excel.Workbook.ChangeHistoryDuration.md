---
title: Workbook.ChangeHistoryDuration property (Excel)
keywords: vbaxl10.chm199080
f1_keywords:
- vbaxl10.chm199080
api_name:
- Excel.Workbook.ChangeHistoryDuration
ms.assetid: 5ebc3cc5-dffa-60cf-08cb-b2f84424c4b4
ms.date: 05/29/2019
ms.localizationpriority: medium
---


# Workbook.ChangeHistoryDuration property (Excel)

Returns or sets the number of days shown in the shared workbook's change history. Read/write **Long**.


## Syntax

_expression_.**ChangeHistoryDuration**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

Any changes in the change history older than the setting for this property are removed when the workbook is closed.


## Example

This example sets the number of days shown in the change history for the active workbook if change tracking is enabled. Any changes in the change history older than the setting for this property are removed when the workbook is closed.

```vb
With ActiveWorkbook 
 If .KeepChangeHistory Then 
 .ChangeHistoryDuration = 7 
 End If 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]