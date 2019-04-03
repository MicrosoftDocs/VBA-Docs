---
title: Results.ResetColumns method (Outlook)
keywords: vbaol11.chm509
f1_keywords:
- vbaol11.chm509
ms.prod: outlook
api_name:
- Outlook.Results.ResetColumns
ms.assetid: 1839dd92-cbab-5fac-a59b-b1ceb6ef874a
ms.date: 06/08/2017
localization_priority: Normal
---


# Results.ResetColumns method (Outlook)

Clears the properties that have been cached with the  **[SetColumns](Outlook.Results.SetColumns.md)** method.


## Syntax

_expression_. `ResetColumns`

_expression_ A variable that represents a [Results](Outlook.Results.md) object.


## Remarks

All properties are accessible after calling the  **ResetColumns** method. **SetColumns** should be reused to store new properties again. **ResetColumns** does nothing if **SetColumns** has not been called first.


## See also


[Results Object](Outlook.Results.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]