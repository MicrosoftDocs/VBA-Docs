---
title: CustomXMLParts.Add method (Office)
keywords: vbaof11.chm298004
f1_keywords:
- vbaof11.chm298004
ms.prod: office
api_name:
- Office.CustomXMLParts.Add
ms.assetid: f2c1588b-c11b-49ca-5db6-4fa4c26d10c5
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLParts.Add method (Office)

Allows you to add a new **CustomXMLPart** to a file.


## Syntax

_expression_.**Add** (_XML_, _SchemaCollection_)

_expression_ An expression that returns a **[CustomXMLParts](Office.CustomXMLParts.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XML_|Optional|**String**|Contains the XML to add to the newly created **CustomXMLPart**.|
| _SchemaCollection_|Optional|**CustomXMLSchemaCollection**|Represents the set of schemas to be used to validate this stream.|

## Return value

CustomXMLPart


## Example

The following example adds a new **CustomXMLPart**, selects a **CustomXMLPart** by using a search criteria, and then selects a single node from that part.


```vb
Sub ShowCustomXmlParts() 
    On Error GoTo Err 
 
    Dim cxp1 As CustomXMLPart 
 
    Dim cxn As CustomXMLNode 
    Dim cxns As CustomXMLNodes 
    Dim strXml As String 
    Dim strUri As String 
 
        ' Example written for Word. 
 
        ' Add a custom XML part. 
        ActiveDocument.CustomXMLParts.Add "<custXMLPart />" 
 
        ' Returns the first custom XML part with the given root namespace. 
        Set cxp1 = ActiveDocument.CustomXMLParts("urn:invoice:namespace")         
 
        ' Get a node using XPath.                              
        Set cxn = cxp1.SelectSingleNode("//*[@quantity < 4]")  
     
    Exit Sub 
                 
' Exception handling. Show the message and resume. 
Err: 
        MsgBox (Err.Description) 
        Resume Next 
End Sub
```


## See also

- [CustomXMLParts object members](overview/library-reference/customxmlparts-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]