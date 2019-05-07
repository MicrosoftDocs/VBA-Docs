---
title: Options.StoreRSIDOnSave property (Word)
keywords: vbawd10.chm162988479
f1_keywords:
- vbawd10.chm162988479
ms.prod: word
api_name:
- Word.Options.StoreRSIDOnSave
ms.assetid: 6f50f3c8-f775-d9d3-2cab-b1c99abf1756
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.StoreRSIDOnSave property (Word)

 **True** for Microsoft Word to assign a random number to changes in a document, each time a document is saved, to facilitate comparing and merging documents. Read/write **Boolean**.


## Syntax

_expression_. `StoreRSIDOnSave`

_expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Remarks

Word stores the random numbers in a table and updates the table after each save. The default for the  **StoreRSIDOnSave** property is **True**. However, RSID information is not saved for HTML documents.

Use the  **[RemovePersonalInformation](Word.Document.RemovePersonalInformation.md)** property if you want to remove information related to authors and reviewers of a document.


## Example

This example turns off storing a random number when saving documents.


```vb
Sub SaveRandomNumber() 
 Application.Options.StoreRSIDOnSave = False 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]