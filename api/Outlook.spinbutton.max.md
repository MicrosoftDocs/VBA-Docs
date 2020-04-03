---
title: SpinButton.Max Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: f8f77453-cc53-68c2-6574-bb2c665e1b76
ms.date: 06/08/2017
localization_priority: Normal
---


# SpinButton.Max Property (Outlook Forms Script)

Returns or sets a **Long** that specifies the maximum and minimum acceptable values for the **[Value](Outlook.spinbutton.value.md)** property of a **[SpinButton](Outlook.spinbutton.md)**. Read/write.


## Syntax

_expression_.**Max**

_expression_ A variable that represents a **SpinButton** object.


## Remarks

Clicking a **SpinButton** changes the **Value** property of the control.

Any integer is an acceptable setting for this property. The recommended range of values is from -32,767 to +32,767. The default value is 1.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]