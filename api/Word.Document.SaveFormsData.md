---
title: Document.SaveFormsData property (Word)
keywords: vbawd10.chm158007347
f1_keywords:
- vbawd10.chm158007347
ms.prod: word
api_name:
- Word.Document.SaveFormsData
ms.assetid: 0f8a14be-49e9-06d4-d601-aa724c4c3c42
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SaveFormsData property (Word)

 **True** if Microsoft Word saves the data entered in a form as a tab-delimited record for use in a database. Read/write **Boolean**.


## Syntax

_expression_. `SaveFormsData`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example sets Word to save only the data entered in a form


```vb
ActiveDocument.SaveFormsData = True
```

This example returns the current status of the  **Save data only for forms** check box in the **Save** options area on the **Save** tab in the **Options** dialog box.




```vb
temp = ActiveDocument.SaveFormsData
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]