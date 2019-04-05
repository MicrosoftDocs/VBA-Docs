---
title: Application.WindowActivate event (Excel)
keywords: vbaxl10.chm504091
f1_keywords:
- vbaxl10.chm504091
ms.prod: excel
api_name:
- Excel.Application.WindowActivate
ms.assetid: 5c618983-27d8-49b1-0a52-001c7a1f94d8
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WindowActivate event (Excel)

Occurs when any workbook window is activated.


## Syntax

_expression_.**WindowActivate** (_Wb_, _Wn_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**| The workbook displayed in the activated window.|
| _Wn_|Required| **[Window](Excel.Window.md)**| The activated window.|

## Remarks

For information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]