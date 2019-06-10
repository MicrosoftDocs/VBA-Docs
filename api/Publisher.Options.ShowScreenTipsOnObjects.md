---
title: Options.ShowScreenTipsOnObjects property (Publisher)
keywords: vbapb10.chm1048608
f1_keywords:
- vbapb10.chm1048608
ms.prod: publisher
api_name:
- Publisher.Options.ShowScreenTipsOnObjects
ms.assetid: b5503200-31fd-72ac-de28-ace55a7123b3
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.ShowScreenTipsOnObjects property (Publisher)

**True** for Microsoft Publisher to display ScreenTips when the mouse pointer hovers over a text box, shape, or other object. Read/write **Boolean**.


## Syntax

_expression_.**ShowScreenTipsOnObjects**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Boolean


## Example

This example disables displaying ScreenTips on objects.

```vb
Sub DisableScreenTips() 
 Options.ShowScreenTipsOnObjects = False 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]