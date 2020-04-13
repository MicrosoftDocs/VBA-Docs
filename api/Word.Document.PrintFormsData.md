---
title: Document.PrintFormsData property (Word)
keywords: vbawd10.chm158007379
f1_keywords:
- vbawd10.chm158007379
ms.prod: word
api_name:
- Word.Document.PrintFormsData
ms.assetid: d4582018-b119-a7a3-27c4-cf4f35d00c19
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.PrintFormsData property (Word)

 **True** if Microsoft Word prints onto a preprinted form only the data entered in the corresponding online form. Read/write **Boolean**.


## Syntax

_expression_. `PrintFormsData`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example sets Word to print only the data from an online form, and then it prints the active document.


```vb
ActiveDocument.PrintFormsData = True 
ActiveDocument.PrintOut
```

This example returns the current status of the **Print data only for forms** check box in the **Options for current document only** area on the **Print** tab in the **Options** dialog box.




```vb
temp = ActiveDocument.PrintFormsData
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]