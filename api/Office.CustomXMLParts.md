---
title: CustomXMLParts object (Office)
keywords: vbaof11.chm300000
f1_keywords:
- vbaof11.chm300000
ms.prod: office
api_name:
- Office.CustomXMLParts
ms.assetid: 98c1c58e-a08d-6304-8626-1e6705917da3
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLParts object (Office)

Represents a collection of **[CustomXMLPart](Office.CustomXMLPart.md)** objects.


## Remarks

There are three default parts that are always created with a document. These are cover pages, doc properties, and app properties. The last two were in previous versions of Microsoft Word but are now provided in XML form in the **CustomXMLParts** object collection.


## Example

The following example adds a node to a **CustomXMLPart** object that is part of the **CustomXMLParts** object collection.


```vb
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```


## See also

- [CustomXMLParts object members](overview/library-reference/customxmlparts-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]