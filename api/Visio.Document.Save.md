---
title: Document.Save method (Visio)
keywords: vis_sdr.chm10516495
f1_keywords:
- vis_sdr.chm10516495
ms.prod: visio
api_name:
- Visio.Document.Save
ms.assetid: 5a9f104c-4893-c401-0093-bc860adf9a4b
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Save method (Visio)

Saves a document.


## Syntax

_expression_.**Save**

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Return value

Integer


## Remarks

To save and name a new document, use the **SaveAs** method. Until a document has been saved, the **Save** method generates an error.


## Example

The following macro shows how to save a Microsoft Visio document.


```vb
Public Sub Save_Example() 
 
 ThisDocument.Save 
 Debug.Print "Document saved." 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]