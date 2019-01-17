---
title: CustomXMLNodes object (Office)
keywords: vbaof11.chm293000
f1_keywords:
- vbaof11.chm293000
ms.prod: office
api_name:
- Office.CustomXMLNodes
ms.assetid: 7aa5b7ae-7d4e-4b57-23b5-b027f39e5ff6
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLNodes object (Office)

Contains a collection of **[CustomXMLNode](Office.CustomXMLNode.md)** objects representing the XML nodes in a document.


## Remarks

The **[Attributes](office.customxmlnode.attributes.md)** and the **[ChildNodes](office.customxmlnode.childnodes.md)** properties return collections of nodes of this type.


## Example

The following example selects one or more nodes matching the XPath expression.


```vb
Sub CustomXmlNodes() 
    Dim cxp1 As CustomXMLPart 
    Dim cxns As CustomXMLNodes 
 
    With ActiveDocument 
  
        ' Returns the first custom xml part with the given root namespace. 
        Set cxp1 = .CustomXMLParts("urn:invoice:namespace")  
         
        ' Get custom xml nodes using XPath.                              
        Set cxns = cxp1.SelectNodes("//*[@unitPrice > 20]")  
                      
    End With 
     
End Sub 

```


## See also

- [CustomXMLNodes object members](overview/library-reference/customxmlnodes-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]