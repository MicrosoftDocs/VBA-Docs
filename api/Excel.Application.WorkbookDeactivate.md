---
title: Application.WorkbookDeactivate Event (Excel)
keywords: vbaxl10.chm504083
f1_keywords:
- vbaxl10.chm504083
ms.prod: excel
api_name:
- Excel.Application.WorkbookDeactivate
ms.assetid: 0a6a55ea-5374-4de7-e48e-e52d903cc749
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WorkbookDeactivate Event (Excel)

Occurs when any open workbook is deactivated.


## Syntax

_expression_. `WorkbookDeactivate`( `_Wb_` )

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook.|

## Return value

Nothing


## Example

This example arranges all open windows when a workbook is deactivated.


```vb
Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]