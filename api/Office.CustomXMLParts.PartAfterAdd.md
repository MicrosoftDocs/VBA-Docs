---
title: CustomXMLParts.PartAfterAdd event (Office)
keywords: vbaof11.chm299001
f1_keywords:
- vbaof11.chm299001
ms.prod: office
api_name:
- Office.CustomXMLParts.PartAfterAdd
ms.assetid: c1a263a5-94cb-f563-145b-151a52a31d52
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLParts.PartAfterAdd event (Office)

Occurs just after a **CustomXMLPart** object is added to the **CustomXMLParts** collection.


## Syntax

_expression_.**PartAfterAdd**(_NewPart_)

_expression_ An expression that returns a **[CustomXMLParts](Office.CustomXMLParts.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NewPart_|Required|**CustomXMLPart**|The part that was added.|

## Example

The following example displays the XML contents of a part after it has been added to a **CustomXMLParts** collection.


```vb
Sub CustomXMLParts_PartAfterAdd(ByVal objPart As CustomXMLPart) 
Dim strPartXML As String 
strPartXML = objPart.XML 
   MsgBox ("The part's contents are: " & vbCrLf & strPartXML) 
End Sub
```


## See also

- [CustomXMLParts object members](overview/library-reference/customxmlparts-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]