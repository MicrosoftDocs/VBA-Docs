---
title: Options.ShowTipPages property (Publisher)
keywords: vbapb10.chm1048609
f1_keywords:
- vbapb10.chm1048609
ms.prod: publisher
api_name:
- Publisher.Options.ShowTipPages
ms.assetid: 44f91cf1-68e3-0755-3114-5dc41a2e4eba
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.ShowTipPages property (Publisher)

**True** for Microsoft Publisher to display tippages in balloons. Read/write **Boolean**.


## Syntax

_expression_.**ShowTipPages**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Boolean


## Example

This example disables displaying tippages in balloons.

```vb
Sub DontShowTipPages() 
 Options.ShowTipPages = False 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]