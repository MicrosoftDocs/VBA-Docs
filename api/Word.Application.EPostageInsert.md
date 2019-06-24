---
title: Application.EPostageInsert event (Word)
keywords: vbawd10.chm4000015
f1_keywords:
- vbawd10.chm4000015
ms.prod: word
api_name:
- Word.Application.EPostageInsert
ms.assetid: 33dfd513-7782-0877-7bf9-1b23cf995d4b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EPostageInsert event (Word)

Occurs when a user inserts electronic postage into a document.


## Syntax

_expression_.**EPostageInsert** (_Doc_)

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 

For information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The name of the document to which to add electronic postage.|

## Example

This example displays a message when electronic postage is inserted into a document.


```vb
Private Sub AppWord_EPostageInsert(ByVal Doc As Document) 
 MsgBox "You just inserted electronic postage into your document." 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]