---
title: NavigationButton.AutoRepeat property (Access)
keywords: vbaac10.chm10457
f1_keywords:
- vbaac10.chm10457
ms.prod: access
api_name:
- Access.NavigationButton.AutoRepeat
ms.assetid: 2348359a-3d49-7ad1-2afe-21e2b73c4923
ms.date: 03/05/2019
localization_priority: Normal
---


# NavigationButton.AutoRepeat property (Access)

You can use the **AutoRepeat** property to specify whether an event procedure or macro runs repeatedly while a command button on a form remains pressed in. Read/write **Boolean**.


## Syntax

_expression_.**AutoRepeat**

_expression_ A variable that represents a **[NavigationButton](Access.NavigationButton.md)** object.


## Remarks

The default value is **False**.

The initial repeat of the event procedure or macro occurs 0.5 seconds after it first runs. Subsequent repeats occur either 0.25 seconds apart or for the duration of the event procedure or macro, whichever is longer.

If the code attached to the command button causes the current record to change, the **AutoRepeat** property has no effect.

If the code attached to the command button causes changes to another control on a form, use the **[DoEvents](../language/reference/user-interface-help/doevents-function.md)** function to ensure proper screen updating.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]