---
title: Application.WorkbookAddinUninstall event (Excel)
keywords: vbaxl10.chm504089
f1_keywords:
- vbaxl10.chm504089
ms.prod: excel
api_name:
- Excel.Application.WorkbookAddinUninstall
ms.assetid: 8c02eb17-e966-703d-36ed-30ce43a56275
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookAddinUninstall event (Excel)

Occurs when any add-in workbook is uninstalled.


## Syntax

_expression_.**WorkbookAddinUninstall** (_Wb_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The uninstalled workbook.|

## Return value

Nothing

## Remarks

For information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


## Example

This example minimizes the Microsoft Excel window when a workbook is uninstalled as an add-in.

```vb
Private Sub App_WorkbookAddinUninstall(ByVal Wb As Workbook) 
 Application.WindowState = xlMinimized 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]