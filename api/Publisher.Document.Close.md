---
title: Document.Close method (Publisher)
keywords: vbapb10.chm196680
f1_keywords:
- vbapb10.chm196680
ms.prod: publisher
api_name:
- Publisher.Document.Close
ms.assetid: b4b21484-1858-b7b3-291f-18ef8cab8ba7
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.Close method (Publisher)

Closes the current publication and creates a blank publication in its place.


## Syntax

_expression_.**Close**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Remarks

You can use the **Close** method only on an open **Document** object in another instance of Microsoft Publisher. Attempting to close the active publication in the current instance of Publisher causes an error.


## Example

This example opens a publication in a new instance of Publisher for modification and then closes the publication. Note that to make this example work, you must replace `FileName` with a valid file name.

```vb
Sub ModifyAnotherPublication() 
 
 ' Create new instance of Publisher. 
 Dim appPub As New Publisher.Application 
 
 ' Open publication. 
 appPub.Open FileName:="FileName" 
 
 ' Put code here to modify the publication as necessary. 
 
 ' Close the publication. 
 appPub.ActiveDocument.Close 
 
 ' Release the other instance of Publisher. 
 Set appPub = Nothing 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]