---
title: CustomXMLNodes object (Office)
keywords: vbaof11.chm293000
f1_keywords:
- vbaof11.chm293000
ms.prod: office
api_name:
- Office.CustomXMLNodes
ms.assetid: 7aa5b7ae-7d4e-4b57-23b5-b027f39e5ff6
ms.date: 06/08/2017
---


# CustomXMLNodes object (Office)

Contains a collection of  **CustomXMLNodes** objects representing the XML nodes in a document.


## Remarks

The  **Attributes** and the **ChildNodes** properties return collections of nodes of this type.


## Example

The following example selects one or more matching the XPath expression.


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


## Properties



|**Name**|
|:-----|
|[Application](Office.CustomXMLNodes.Application.md)|
|[Count](Office.CustomXMLNodes.Count.md)|
|[Creator](Office.CustomXMLNodes.Creator.md)|
|[Item](Office.CustomXMLNodes.Item.md)|
|[Parent](Office.CustomXMLNodes.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
