---
title: Document.BeforeClose event (Publisher)
keywords: vbapb10.chm285212674
f1_keywords:
- vbapb10.chm285212674
ms.prod: publisher
api_name:
- Publisher.Document.BeforeClose
ms.assetid: d40e36b6-fea7-a9d5-0c88-55197983b888
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.BeforeClose event (Publisher)

Occurs immediately before any open document closes.


## Syntax

_expression_.**BeforeClose** (_Cancel_)

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the document doesn't close when the procedure is finished.|

## Remarks

For more information about using events with the **Document** object, see [Using events with the Document object](../publisher/Concepts/using-events-with-the-document-object-publisher.md).


## Example

This example prompts the user for a yes or no response before closing a document. For this example to work, you must place this code into the **ThisDocument** module.

```vb
Private Sub Document_BeforeClose(Cancel As Boolean) 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really want to close " _ 
 & "the document?", vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]