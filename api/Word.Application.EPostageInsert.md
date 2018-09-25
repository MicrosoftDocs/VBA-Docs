---
title: Application.EPostageInsert Event (Word)
keywords: vbawd10.chm4000015
f1_keywords:
- vbawd10.chm4000015
ms.prod: word
api_name:
- Word.Application.EPostageInsert
ms.assetid: 33dfd513-7782-0877-7bf9-1b23cf995d4b
ms.date: 06/08/2017
---


# Application.EPostageInsert Event (Word)

Occurs when a user inserts electronic postage into a document.


## Syntax

 _expression_. `Private Sub object_EPostageInsert`( `_ByVal Doc As Document_` )

 _expression_ A variable that represents an '[Application](Word.Application.md)' object that has been declared with events in a class module. For information about using events with the **Application** object, see [Using Events with the Application Object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
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

