---
title: ExchangeUser.Details method (Outlook)
keywords: vbaol11.chm2074
f1_keywords:
- vbaol11.chm2074
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Details
ms.assetid: 6c93a583-cc61-e527-7832-88dba525854a
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.Details method (Outlook)

Displays a modal dialog box that provides detailed information about an  **[ExchangeUser](Outlook.ExchangeUser.md)** object.


## Syntax

_expression_. `Details`( `_HWnd_` )

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HWnd_|Optional| **Variant**| The parent window handle for the Details dialog box. A zero value (the default) specifies a modal dialog box.|

## Remarks

The  **Details** method fails if the **[ExchangeUser.Name](Outlook.ExchangeUser.Name.md)** property is empty. You must use error handling to handle run-time errors, and when the user clicks **Cancel** in the dialog box.

The  **Details** method actually stops the code from running while the dialog box is displayed.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]