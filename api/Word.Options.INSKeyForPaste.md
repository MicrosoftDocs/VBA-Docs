---
title: Options.INSKeyForPaste property (Word)
keywords: vbawd10.chm162988102
f1_keywords:
- vbawd10.chm162988102
ms.prod: word
api_name:
- Word.Options.INSKeyForPaste
ms.assetid: a16b57f1-8c56-9544-4da2-57a114f14081
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.INSKeyForPaste property (Word)

 **True** if the INS key can be used for pasting the Clipboard contents. Read/write **Boolean**.


## Syntax

_expression_. `INSKeyForPaste`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example enables the INS key to be used for pasting the contents of the Clipboard.


```vb
Options.INSKeyForPaste = True
```

This example returns the status of the Use the INS key for paste option on the Edit tab in the Options dialog box.




```vb
Dim blnTemp As Boolean 
 
blnTemp = Options.INSKeyForPaste
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]