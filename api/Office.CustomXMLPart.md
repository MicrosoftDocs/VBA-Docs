---
title: CustomXMLPart Object (Office)
keywords: vbaof11.chm297000
f1_keywords:
- vbaof11.chm297000
ms.prod: office
api_name:
- Office.CustomXMLPart
ms.assetid: a4f90bac-01d6-bba4-f64b-a64e2b122cfd
ms.date: 06/08/2017
---


# CustomXMLPart Object (Office)

Represents a single  **CustomXMLPart** in a **CustomXMLParts** collection.


## Example

The following example adds a part to a  **CustomXMLPart** object.


```vb
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```


## See also


[Object Model Reference](overview/reference-object-library-reference-for-office.md)
#### Other resources


[CustomXMLPart Object Members](./overview/customxmlpart-members-office.md)

