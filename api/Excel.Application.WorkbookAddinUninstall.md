---
title: Application.WorkbookAddinUninstall Event (Excel)
keywords: vbaxl10.chm504089
f1_keywords:
- vbaxl10.chm504089
ms.prod: excel
api_name:
- Excel.Application.WorkbookAddinUninstall
ms.assetid: 8c02eb17-e966-703d-36ed-30ce43a56275
ms.date: 06/08/2017
---


# Application.WorkbookAddinUninstall Event (Excel)

Occurs when any add-in workbook is uninstalled.


## Syntax

 _expression_. `WorkbookAddinUninstall`( `_Wb_` )

 _expression_ A variable that represents an [Application](Excel.Application(Graph property).md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The uninstalled workbook.|

### Return Value

Nothing


## Example

This example minimizes the Microsoft Excel window when a workbook is installed as an add-in.


```vb
Private Sub App_WorkbookAddinUninstall(ByVal Wb As Workbook) 
 Application.WindowState = xlMinimized 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)

