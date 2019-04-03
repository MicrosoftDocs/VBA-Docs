---
title: Items.ResetColumns method (Outlook)
keywords: vbaol11.chm69
f1_keywords:
- vbaol11.chm69
ms.prod: outlook
api_name:
- Outlook.Items.ResetColumns
ms.assetid: 0543dd17-1e65-5484-ab21-d4791b3b1194
ms.date: 06/08/2017
localization_priority: Normal
---


# Items.ResetColumns method (Outlook)

Clears the properties that have been cached with the  **[SetColumns](Outlook.Items.SetColumns.md)** method.


## Syntax

_expression_. `ResetColumns`

_expression_ A variable that represents an [Items](Outlook.Items.md) object.


## Remarks

All properties are accessible after calling the  **ResetColumns** method. **SetColumns** should be reused to store new properties again. **ResetColumns** does nothing if **SetColumns** has not been called first.


## See also


[Items Object](Outlook.Items.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]