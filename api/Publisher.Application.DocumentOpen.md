---
title: Application.DocumentOpen event (Publisher)
keywords: vbapb10.chm268435463
f1_keywords:
- vbapb10.chm268435463
ms.prod: publisher
api_name:
- Publisher.Application.DocumentOpen
ms.assetid: 3bdd4b38-ec40-a08f-3742-f81a6ed333b3
ms.date: 06/04/2019
localization_priority: Normal
---


# Application.DocumentOpen event (Publisher)

Occurs when opening a document.


## Syntax

_expression_.**DocumentOpen** (_Doc_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Doc_|Required| **Document**|The document that's being opened.|

## Example

This example displays a message with the document's name when opening a document.

```vb
Private Sub appPub_DocumentOpen(ByVal Doc As Document) 
 MsgBox "Please wait. " & Doc.Name & " is opening." 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]