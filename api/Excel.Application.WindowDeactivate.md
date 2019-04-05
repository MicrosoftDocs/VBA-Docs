---
title: Application.WindowDeactivate event (Excel)
keywords: vbaxl10.chm504092
f1_keywords:
- vbaxl10.chm504092
ms.prod: excel
api_name:
- Excel.Application.WindowDeactivate
ms.assetid: 6adcba54-3d4a-f780-915e-5798303faf60
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WindowDeactivate event (Excel)

Occurs when any workbook window is deactivated.


## Syntax

_expression_.**WindowDeactivate** (_Wb_, _Wn_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook displayed in the deactivated window.|
| _Wn_|Required| **[Window](Excel.Window.md)**|The deactivated window.|

## Remarks

For information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


## Example

This example minimizes any workbook window when it's deactivated.

```vb
Private Sub Workbook_WindowDeactivate(ByVal Wn As Excel.Window) 
 Wn.WindowState = xlMinimized 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]