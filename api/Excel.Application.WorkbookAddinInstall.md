---
title: Application.WorkbookAddinInstall event (Excel)
keywords: vbaxl10.chm504088
f1_keywords:
- vbaxl10.chm504088
ms.prod: excel
api_name:
- Excel.Application.WorkbookAddinInstall
ms.assetid: 955c8f2a-4647-ed7e-29f9-8d6d165898ec
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookAddinInstall event (Excel)

Occurs when a workbook is installed as an add-in.


## Syntax

_expression_.**WorkbookAddinInstall** (_Wb_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The installed workbook.|

## Return value

Nothing

## Remarks

For information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


## Example

This example maximizes the Microsoft Excel window when a workbook is installed as an add-in.

```vb
Private Sub App_WorkbookAddinInstall(ByVal Wb As Workbook) 
 Application.WindowState = xlMaximized 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]