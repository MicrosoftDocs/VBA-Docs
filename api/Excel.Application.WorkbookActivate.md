---
title: Application.WorkbookActivate event (Excel)
keywords: vbaxl10.chm504082
f1_keywords:
- vbaxl10.chm504082
ms.prod: excel
api_name:
- Excel.Application.WorkbookActivate
ms.assetid: a2b6ea2e-3753-69bf-9a81-ec2fce29d4fd
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookActivate event (Excel)

Occurs when any workbook is activated.


## Syntax

_expression_.**WorkbookActivate** (_Wb_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The activated workbook.|

## Return value

Nothing

## Remarks

For information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


## Example

This example arranges open windows when a workbook is activated.

```vb
Private Sub App_WorkbookActivate(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]