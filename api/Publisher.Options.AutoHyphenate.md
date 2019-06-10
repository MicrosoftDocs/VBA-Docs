---
title: Options.AutoHyphenate property (Publisher)
keywords: vbapb10.chm1048580
f1_keywords:
- vbapb10.chm1048580
ms.prod: publisher
api_name:
- Publisher.Options.AutoHyphenate
ms.assetid: 821d0540-80ec-9f9d-777e-4d2596baf7d7
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.AutoHyphenate property (Publisher)

**True** (default) for Microsoft Publisher to automatically hyphenate text in text frames. Read/write **Boolean**.


## Syntax

_expression_.**AutoHyphenate**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Boolean


## Example

This example turns on automatic hyphenation for Publisher and sets the amount of space from the right margin to use when hyphenating words to one inch (72 points).

```vb
Sub SetHyphenationZone() 
 With Options 
 .AutoHyphenate = True 
 .HyphenationZone = 72 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]