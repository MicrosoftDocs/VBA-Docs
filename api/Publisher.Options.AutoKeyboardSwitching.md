---
title: Options.AutoKeyboardSwitching property (Publisher)
keywords: vbapb10.chm1048627
f1_keywords:
- vbapb10.chm1048627
ms.prod: publisher
api_name:
- Publisher.Options.AutoKeyboardSwitching
ms.assetid: 05f22aa6-332d-e033-ab9d-550eb08f1018
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.AutoKeyboardSwitching property (Publisher)

**True** for Microsoft Publisher to automatically switch the keyboard language to the language used for the text at the cursor position. Read/write **Boolean**.


## Syntax

_expression_.**AutoKeyboardSwitching**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Boolean


## Example

This example enables automatically switching the keyboard language to the necessary language.

```vb
Sub SetGlobalOptions() 
 Options.AutoKeyboardSwitching = True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]