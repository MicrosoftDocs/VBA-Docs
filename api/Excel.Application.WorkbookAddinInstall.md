---
title: Application.WorkbookAddinInstall Event (Excel)
keywords: vbaxl10.chm504088
f1_keywords:
- vbaxl10.chm504088
ms.prod: excel
api_name:
- Excel.Application.WorkbookAddinInstall
ms.assetid: 955c8f2a-4647-ed7e-29f9-8d6d165898ec
ms.date: 06/08/2017
---


# Application.WorkbookAddinInstall Event (Excel)

Occurs when a workbook is installed as an add-in.


## Syntax

 _expression_. `WorkbookAddinInstall`( `_Wb_` )

 _expression_ A variable that represents an [Application](Excel.Application(Graph property).md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The installed workbook.|

### Return value

Nothing


## Example

This example maximizes the Microsoft Excel window when a workbook is installed as an add-in.


```vb
Private Sub App_WorkbookAddinInstall(ByVal Wb As Workbook) 
 Application.WindowState = xlMaximized 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)

