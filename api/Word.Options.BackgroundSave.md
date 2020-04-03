---
title: Options.BackgroundSave property (Word)
keywords: vbawd10.chm162988078
f1_keywords:
- vbawd10.chm162988078
ms.prod: word
api_name:
- Word.Options.BackgroundSave
ms.assetid: a579d9ae-5ee2-543e-fe16-e642e48dcb61
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.BackgroundSave property (Word)

 **True** if Word saves documents in the background. When Word is saving in the background, users can continue to type and to choose commands. Read/write **Boolean**.


## Syntax

_expression_. `BackgroundSave`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Example

This example allows users to continue working in a document while Word is saving it.


```vb
Options.BackgroundSave = True
```

This example returns the current status of the **Allow background saves** option on the **Save** tab in the **Options** dialog box.




```vb
Dim blnAutoSave As Boolean 
 
blnAutoSave = Options.BackgroundSave
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]