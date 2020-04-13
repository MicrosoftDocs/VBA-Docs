---
title: Document.SetLetterContent method (Word)
keywords: vbawd10.chm158007421
f1_keywords:
- vbawd10.chm158007421
ms.prod: word
api_name:
- Word.Document.SetLetterContent
ms.assetid: 8c9b2f6e-34a7-41a3-761d-c1a5da141aba
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SetLetterContent method (Word)

Inserts the contents of the specified  **LetterContent** object into a document.


## Syntax

_expression_. `SetLetterContent`( `_LetterContent_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _LetterContent_|Required| **[LetterContent](Word.LetterContent.md)**|The that includes the various elements of the letter.|

## Remarks

This method is similar to the **RunLetterWizard** method except that it doesn't display the Letter Wizard dialog box. The method adds, deletes, or restyles letter elements in the specified document based on the contents of the **LetterContent** object.


## Example

This example retrieves the Letter Wizard elements from the active document, changes the attention line text, and then uses the **SetLetterContent** method to update the active document to reflect the changes.


```vb
Set myLetterContent = ActiveDocument.GetLetterContent 
myLetterContent.AttentionLine = "Greetings" 
ActiveDocument.SetLetterContent LetterContent:=myLetterContent
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]