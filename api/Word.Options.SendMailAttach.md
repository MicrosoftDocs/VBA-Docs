---
title: Options.SendMailAttach property (Word)
keywords: vbawd10.chm162988056
f1_keywords:
- vbawd10.chm162988056
ms.prod: word
api_name:
- Word.Options.SendMailAttach
ms.assetid: e749ca30-089f-5116-ce70-a3d760006a2c
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.SendMailAttach property (Word)

 **True** if the **Send To** command on the **File** menu inserts the active document as an attachment to a mail message. Read/write **Boolean**.


## Syntax

_expression_. `SendMailAttach`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

 **False** if the **Send To** command inserts the contents of the active document as text in a mail message.


## Example

This example opens a new mail message that has the active document attached to it.


```vb
Options.SendMailAttach = True 
ActiveDocument.SendMail
```

This example returns the state of the  **Mail as attachment** option on the **General** tab of the **Options** dialog box.




```vb
Msgbox Options.SendMailAttach
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]