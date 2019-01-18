---
title: DocumentProperty.Parent property (Office)
keywords: vbaof11.chm250003
f1_keywords:
- vbaof11.chm250003
ms.prod: office
api_name:
- Office.DocumentProperty.Parent
ms.assetid: 4d6e4c41-09d2-7e0b-c35b-fde629c53c46
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperty.Parent property (Office)

Gets the **Parent** object for the **DocumentProperty** object. Read-only.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a **[DocumentProperty](Office.DocumentProperty.md)** object.


## Return value

Object


## Example

This example displays the name of the parent object for a document property. You must pass a valid **DocumentProperty** object to the procedure.


```vb
Sub DisplayParent(dp as DocumentProperty) 
 MsgBox dp.Parent.Name 
End Sub
```


## See also

- [DocumentProperty object members](overview/library-reference/documentproperty-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]