---
title: Workbook.ConflictResolution property (Excel)
keywords: vbaxl10.chm199091
f1_keywords:
- vbaxl10.chm199091
ms.prod: excel
api_name:
- Excel.Workbook.ConflictResolution
ms.assetid: 5142c848-0731-14d9-5913-bbaa67bf308f
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ConflictResolution property (Excel)

Returns or sets the way conflicts are to be resolved whenever a shared workbook is updated. Read/write **[XlSaveConflictResolution](Excel.XlSaveConflictResolution.md)**.


## Syntax

_expression_.**ConflictResolution**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example causes the local user's changes to be accepted whenever there's a conflict in the shared workbook.

```vb
ActiveWorkbook.ConflictResolution = xlLocalSessionChanges 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]