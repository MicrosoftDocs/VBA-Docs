---
title: CustomXMLPart.LoadXML method (Office)
keywords: vbaof11.chm295011
f1_keywords:
- vbaof11.chm295011
ms.prod: office
api_name:
- Office.CustomXMLPart.LoadXML
ms.assetid: efdbb098-48ec-1c64-9d9d-b0a64a5c3753
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLPart.LoadXML method (Office)

Allows the template author to populate a **CustomXMLPart** object from an XML string. Returns **True** if the load was successful.


## Syntax

_expression_.**LoadXML**(_XML_)

_expression_ An expression that returns a **[CustomXMLPart](Office.CustomXMLPart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XML_|Required|**String**|Contains the XML to load.|

## Return value

Boolean


## Example

The following example loads XML into a custom XML part from a string.


```vb
Sub ShowCustomXmlParts() 
    On Error GoTo Err 
 
    Dim cxp1 As CustomXMLPart 
 
        ' Add a custom XML part and then load the XML. 
        Set cxp1 = ActiveDocument.CustomXMLParts.Add 
        cxp1.LoadXML("<discounts><discount>0.10</discount></discounts>") 
     
    Exit Sub 
                 
' Exception handling. Show the message and resume. 
Err: 
        MsgBox (Err.Description) 
        Resume Next 
End Sub
```


## See also

- [CustomXMLPart object members](overview/library-reference/customxmlpart-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]