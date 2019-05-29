---
title: Workbook.UpdateFromFile method (Excel)
keywords: vbaxl10.chm199159
f1_keywords:
- vbaxl10.chm199159
ms.prod: excel
api_name:
- Excel.Workbook.UpdateFromFile
ms.assetid: f5148b60-9b25-8a12-5cf3-40103dcff2a3
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.UpdateFromFile method (Excel)

Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory. If the disk copy hasn't changed since the workbook was loaded, the in-memory copy of the workbook isn't reloaded.


## Syntax

_expression_.**UpdateFromFile**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

This method is useful when a workbook is opened as read-only by user A and opened as read/write by user B. If user B saves a newer version of the workbook to disk while user A still has the workbook open, user A cannot get the updated copy without closing and reopening the workbook and losing view settings. The **UpdateFromFile** method updates the in-memory copy of the workbook from the disk file.


## Example

This example updates the active workbook from the disk version of the file.

```vb
ActiveWorkbook.UpdateFromFile
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]