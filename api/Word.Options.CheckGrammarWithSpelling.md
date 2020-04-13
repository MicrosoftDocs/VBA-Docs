---
title: Options.CheckGrammarWithSpelling property (Word)
keywords: vbawd10.chm162988317
f1_keywords:
- vbawd10.chm162988317
ms.prod: word
api_name:
- Word.Options.CheckGrammarWithSpelling
ms.assetid: b08d1bc4-bc9c-c9b3-0448-92a051809a25
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.CheckGrammarWithSpelling property (Word)

 **True** if Word checks grammar while checking spelling. Read/write **Boolean**.


## Syntax

_expression_. `CheckGrammarWithSpelling`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Remarks

This property controls whether Word checks grammar when you check spelling by using the **Spelling** command (**Tools** menu).

To check spelling or grammar from a Visual Basic procedure, use the **[CheckSpelling](Word.Application.CheckSpelling.md)** method to check only spelling and use the **[CheckGrammar](Word.Application.CheckGrammar.md)** method to check both grammar and spelling.


## Example

This example returns the status of the **Check grammar with spelling** option on the **Spelling & Grammar** tab in the **Options** dialog box. If the option is selected, the procedure checks both spelling and grammar for the active document; otherwise, only spelling is checked.


```vb
If Options.CheckGrammarWithSpelling = True Then 
 ActiveDocument.CheckGrammar 
Else 
 ActiveDocument.CheckSpelling 
End If
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]