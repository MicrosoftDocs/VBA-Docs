---
title: Options.UpdateFieldsAtPrint property (Word)
keywords: vbawd10.chm162988062
f1_keywords:
- vbawd10.chm162988062
ms.prod: word
api_name:
- Word.Options.UpdateFieldsAtPrint
ms.assetid: 065d63a9-7c07-c351-b18a-44dfa6b59078
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.UpdateFieldsAtPrint property (Word)

 **True** if Microsoft Word updates fields automatically before printing a document. Read/write **Boolean**.


## Syntax

_expression_. `UpdateFieldsAtPrint`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Word to update fields automatically before printing, and then it prints the active document.


```vb
Options.UpdateFieldsAtPrint = True 
ActiveDocument.PrintOut
```

This example returns the current status of the **Update fields** option on the **Print** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.UpdateFieldsAtPrint
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]