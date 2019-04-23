---
title: Options.Pagination property (Word)
keywords: vbawd10.chm162988051
f1_keywords:
- vbawd10.chm162988051
ms.prod: word
api_name:
- Word.Options.Pagination
ms.assetid: 885a621c-a1fd-e428-80a8-c0a7ca904a22
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.Pagination property (Word)

 **True** if Microsoft Word repaginates documents in the background. Read/write **Boolean**.


## Syntax

_expression_. `Pagination`

 _expression_ An expression that returns a '[Options](Word.Options.md)' object.


## Example

This example sets Word to perform background repagination.


```vb
Options.Pagination = True
```

This example returns the current status of the Background repagination option on the General tab in the Options dialog box (Tools menu).




```vb
temp = Options.Pagination
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]