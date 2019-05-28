---
title: Options.AutoFormatAsYouTypeInsertClosings property (Word)
keywords: vbawd10.chm162988335
f1_keywords:
- vbawd10.chm162988335
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeInsertClosings
ms.assetid: 8e51f053-03df-84c3-cd08-d53281602646
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AutoFormatAsYouTypeInsertClosings property (Word)

 **True** for Microsoft Word to automatically insert the corresponding memo closing when the user enters a memo heading. Read/write.


## Syntax

_expression_. `AutoFormatAsYouTypeInsertClosings`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to automatically insert the corresponding memo closing when the user enters a memo heading.


```vb
Sub AutoInsertClosings() 
 Options.AutoFormatAsYouTypeInsertClosings = True 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]