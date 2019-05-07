---
title: Options.PasteMergeLists property (Word)
keywords: vbawd10.chm162988482
f1_keywords:
- vbawd10.chm162988482
ms.prod: word
api_name:
- Word.Options.PasteMergeLists
ms.assetid: 82989419-32c6-6a70-685f-eae11de50cae
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PasteMergeLists property (Word)

 **True** to merge the formatting of pasted lists with surrounding lists. Read/write **Boolean**.


## Syntax

_expression_. `PasteMergeLists`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Example

This example sets Microsoft Word to automatically merge list formatting with surrounding lists if the option has been disabled.


```vb
Sub UseSmartStyle() 
 With Options 
 If .PasteMergeLists = False Then 
 .PasteMergeLists = True 
 End If 
 End With 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]