---
title: DisplayAlerts property (Excel Graph)
keywords: vbagr10.chm65879
f1_keywords:
- vbagr10.chm65879
ms.prod: excel
api_name:
- Excel.DisplayAlerts
ms.assetid: 630e60be-23e3-795b-1ed9-26b791fb7efc
ms.date: 04/10/2019
localization_priority: Normal
---


# DisplayAlerts property (Excel Graph)

**True** if Graph displays certain alerts and messages while a macro is running. Read/write **Boolean**.

## Syntax

_expression_.**DisplayAlerts**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

The default value is **True**. Set this property to **False** if you don't want to be disturbed by prompts and alert messages while a macro is running; any time a message requires a response, Graph chooses the default response.

If you set this property to **False**, Graph doesn't automatically set it back to **True** when your macro stops running. Write your macro so that it always sets this property back to **True** when it stops running.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]