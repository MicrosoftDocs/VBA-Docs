---
title: Document.Open event (Publisher)
keywords: vbapb10.chm285212673
f1_keywords:
- vbapb10.chm285212673
ms.prod: publisher
api_name:
- Publisher.Document.Open
ms.assetid: 43108d1d-d101-8a07-943e-c9b8dbadcbfd
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.Open event (Publisher)

Occurs when a publication is opening.


## Syntax

_expression_.**Open**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Remarks

To access the **Document** object events, declare a **Document** object variable in the General Declarations section of a class module, and then set the variable equal to the **Document** object for which you want to access events.

For more information about using events with the **Document** object, see [Using events with the Document object](../publisher/Concepts/using-events-with-the-document-object-publisher.md).


## Example

This example displays a message when a publication is opened. The procedure can be stored in the **ThisDocument** module of a publication.

```vb
Private Sub Document_Open() 
 MsgBox "This publication is copyrighted." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]