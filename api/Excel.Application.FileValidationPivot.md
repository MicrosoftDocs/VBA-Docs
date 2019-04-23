---
title: Application.FileValidationPivot property (Excel)
keywords: vbaxl10.chm133336
f1_keywords:
- vbaxl10.chm133336
ms.prod: excel
api_name:
- Excel.Application.FileValidationPivot
ms.assetid: 3cf6e177-9dbe-8ee8-3d84-599d7e2221da
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.FileValidationPivot property (Excel)

Returns or sets how Excel will validate the contents of the data caches for PivotTable reports. Read/write.


## Syntax

_expression_.**FileValidationPivot**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Return value

**[XlFileValidationPivotMode](Excel.XlFileValidationPivotMode.md)**


## Remarks

Files that contain data caches that do not validate will be opened in a Protected View window. If you set the **FileValidationPivot** property, that setting will remain in effect for the entire session that the application is open.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]