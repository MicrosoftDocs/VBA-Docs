---
title: CustomXMLNode object (Office)
keywords: vbaof11.chm294000
f1_keywords:
- vbaof11.chm294000
ms.prod: office
api_name:
- Office.CustomXMLNode
ms.assetid: e90213f5-6d62-52d8-3043-2399eaa5aaba
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLNode object (Office)

Represents an XML node in a tree in a document. The **CustomXMLNode** object is a member of the **CustomXMLNodes** collection.


## Remarks

The **CustomXMLNode** object is designed to have functional parity with the **[IXMLDOMNode](https://docs.microsoft.com/previous-versions/windows/desktop/ms765513(v=vs.85))** interface. In addition, it contains an **[XPath](office.customxmlnode.xpath.md)** property, which is a great improvement over the objects provided by MSXML.


## Example

The following example selects a single node from a **CustomXMLPart** object by using an XPath expression and assigns it to a **CustomXMLNode** object.


```vb
Sub CustomXmlNodes()  
    Dim cxp1 As CustomXMLPart 
    Dim cxn As CustomXMLNode 
 
    With ActiveDocument 
 
        ' Returns the first custom xml part with the given root namespace. 
        Set cxp1 = .CustomXMLParts("urn:invoice:namespace")  
         
        ' Get the first node matching the XPath expression.                              
        Set cxn = cxp1.SelectSingleNode("//*[@quantity < 4]") 
                 
    End With 
     
End Sub
```


## See also

- [CustomXMLNode object members](overview/library-reference/customxmlnode-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]