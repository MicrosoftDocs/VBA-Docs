---
title: Options.TypeNReplace property (Publisher)
keywords: vbapb10.chm1048626
f1_keywords:
- vbapb10.chm1048626
ms.prod: publisher
api_name:
- Publisher.Options.TypeNReplace
ms.assetid: 0eb378d2-3554-6a46-8b6b-4a990b4638db
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.TypeNReplace property (Publisher)

**True** for Microsoft Publisher to replace unreadable Asian character clusters resulting from invalid keyboard sequences. Read/write **Boolean**.


## Syntax

_expression_.**TypeNReplace**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Boolean


## Example

This example instructs Publisher to replace unreadable Asian character clusters resulting from invalid keyboard sequences.

```vb
Sub TypeReplace() 
 Options.TypeNReplace = True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]