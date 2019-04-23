---
title: Application.WindowResize event (Excel)
keywords: vbaxl10.chm504090
f1_keywords:
- vbaxl10.chm504090
ms.prod: excel
api_name:
- Excel.Application.WindowResize
ms.assetid: 937c4b8f-3b37-ada7-ee72-0ad4707c2e2b
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WindowResize event (Excel)

Occurs when any workbook window is resized.


## Syntax

_expression_.**WindowResize** (_Wb_, _Wn_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook displayed in the resized window.|
| _Wn_|Required| **[Window](Excel.Window.md)**|The resized window.|

## Remarks

For information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]