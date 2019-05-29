---
title: Workbook.Permission property (Excel)
keywords: vbaxl10.chm199220
f1_keywords:
- vbaxl10.chm199220
ms.prod: excel
api_name:
- Excel.Workbook.Permission
ms.assetid: ef04f56e-a04d-c3d9-fdda-611be7bf9d39
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Permission property (Excel)

Returns a **[Permission](office.permission.md)** object that represents the permission settings in the specified workbook.


## Syntax

_expression_.**Permission**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

The following example returns the permission settings for the active workbook.

```vb
Dim objPermission As Permission 
 
Set objPermission = ActiveWorkbook.Permission
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]