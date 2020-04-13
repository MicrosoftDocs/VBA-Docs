---
title: ExchangeDistributionList.Details method (Outlook)
keywords: vbaol11.chm2121
f1_keywords:
- vbaol11.chm2121
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.Details
ms.assetid: e1d3a324-1a2b-54e2-641a-f7d37aa37358
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeDistributionList.Details method (Outlook)

Displays a modal dialog box that provides detailed information about an **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)** object.


## Syntax

_expression_. `Details`( `_HWnd_` )

_expression_ A variable that represents an [ExchangeDistributionList](Outlook.ExchangeDistributionList.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HWnd_|Optional| **Variant**| The parent window handle for the Details dialog box. A zero value (the default) specifies a modal dialog box.|

## Remarks

The **Details** method fails if the **[ExchangeDistributionList.Name](Outlook.ExchangeDistributionList.Name.md)** property is empty. You must use error handling to handle run-time errors, and when the user clicks **Cancel** in the dialog box.

The **Details** method actually stops the code from running while the dialog box is displayed.


## See also


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]