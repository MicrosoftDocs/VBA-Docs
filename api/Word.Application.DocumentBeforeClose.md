---
title: Application.DocumentBeforeClose event (Word)
keywords: vbawd10.chm400005
f1_keywords:
- vbawd10.chm400005
ms.prod: word
api_name:
- Word.Application.DocumentBeforeClose
ms.assetid: 91c89b29-3110-85d7-c141-d1add3bb57f1
ms.date: 08/20/2018
localization_priority: Normal
---


# Application.DocumentBeforeClose event (Word)

Occurs immediately before any open document closes.

> [!NOTE] 
> If you are working with a document embedded within another document, this event will not occur.

## Syntax

Private Sub  _expression_ 'DocumentBeforeClose** (_Doc As Document_**, **_Cancel As Boolean_**)

_expression_ A variable that represents an [Application](Word.Application.md) object declared with events in a class module.


## Parameters


|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **[Document](Word.Document.md)**|The document that's being closed.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the document doesn't close when the procedure is finished.|

## Remarks

For more information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Example

This example prompts the user for a yes or no response before closing any document. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md) for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_DocumentBeforeClose _ 
        (ByVal Doc As Document, _ 
        Cancel As Boolean) 
 
    Dim intResponse As Integer 
 
    intResponse = MsgBox("Do you really " _ 
        & "want to close the document?", _ 
        vbYesNo) 
 
    If intResponse = vbNo Then Cancel = True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]