---
title: DocumentProperties.Parent property (Office)
keywords: vbaof11.chm250011
f1_keywords:
- vbaof11.chm250011
ms.prod: office
api_name:
- Office.DocumentProperties.Parent
ms.assetid: e1239ffa-b89e-e78f-4009-d576c473d477
ms.date: 06/08/2017
---


# DocumentProperties.Parent property (Office)

Gets the  **Parent** object for the **DocumentProperties** object. Read-only.


## Syntax

 _expression_. `Parent`

 _expression_ A variable that represents a [DocumentProperties](Office.DocumentProperties.md) object.


## Return value

Object


## Example

This example displays the name of the parent object for a document property. You must pass a valid  **DocumentProperty** object to the procedure.


```vb
Sub DisplayParent(dp as DocumentProperty) 
 MsgBox dp.Parent.Name 
End Sub
```


## See also


[DocumentProperties Object](Office.DocumentProperties.md)



[DocumentProperties Object Members](./overview/Library-Reference/documentproperties-members-office.md)

