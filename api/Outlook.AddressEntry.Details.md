---
title: AddressEntry.Details method (Outlook)
keywords: vbaol11.chm2051
f1_keywords:
- vbaol11.chm2051
ms.prod: outlook
api_name:
- Outlook.AddressEntry.Details
ms.assetid: 85457da6-c97a-387d-6c7e-40eb005b25aa
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressEntry.Details method (Outlook)

Displays a modeless dialog box that provides detailed information about an **[AddressEntry](Outlook.AddressEntry.md)** object.


## Syntax

_expression_. `Details`( `_HWnd_` )

 _expression_ An expression that returns a [AddressEntry](Outlook.AddressEntry.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HWnd_|Optional| **Variant**|The parent window handle for the **Details** dialog box. A zero value (the default) specifies that the dialog is parented to Outlook.|

## Remarks


> [!NOTE] 
> The **Details** method fails if the **[Name](Outlook.AddressEntry.Name.md)** property is empty.

You must use error handling to handle run-time errors when the user clicks **Cancel** in the dialog box. The **Details** method actually stops the code from running while the dialog box is displayed.


## See also


[AddressEntry Object](Outlook.AddressEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]