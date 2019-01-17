---
title: CustomXMLPart object (Office)
keywords: vbaof11.chm297000
f1_keywords:
- vbaof11.chm297000
ms.prod: office
api_name:
- Office.CustomXMLPart
ms.assetid: a4f90bac-01d6-bba4-f64b-a64e2b122cfd
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLPart object (Office)

Represents a single **CustomXMLPart** in a **CustomXMLParts** collection.


## Example

The following example adds a part to a **CustomXMLPart** object.


```vb
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```


## See also

- [CustomXMLPart object members](overview/library-reference/customxmlpart-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]