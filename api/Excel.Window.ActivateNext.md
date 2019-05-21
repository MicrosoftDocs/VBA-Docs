---
title: Window.ActivateNext method (Excel)
keywords: vbaxl10.chm356074
f1_keywords:
- vbaxl10.chm356074
ms.prod: excel
api_name:
- Excel.Window.ActivateNext
ms.assetid: eeef1ef2-b1c5-6618-1f66-827bc64e2033
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.ActivateNext method (Excel)

Activates the specified window and then sends it to the back of the window z-order.


## Syntax

_expression_.**ActivateNext**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Return value

Variant


## Example

This example sends the active window to the back of the z-order.

```vb
ActiveWindow.ActivateNext
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]