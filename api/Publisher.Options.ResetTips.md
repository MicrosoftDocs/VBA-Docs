---
title: Options.ResetTips method (Publisher)
keywords: vbapb10.chm1048616
f1_keywords:
- vbapb10.chm1048616
ms.prod: publisher
api_name:
- Publisher.Options.ResetTips
ms.assetid: a119aacc-ba19-f430-e8af-6d84c438ec25
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.ResetTips method (Publisher)

Resets tippages so that a user can view them when using features that have been used before.


## Syntax

_expression_.**ResetTips**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Remarks

The **ResetTips** method is equivalent to choosing **Reset Tips** on the **User Assistance** tab of the **Options** dialog box (**Tools** menu).


## Example

This example resets tip balloons.

```vb
Sub ResetTippages() 
 Options.ResetTips 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]