---
title: CustomXMLParts.PartAfterLoad Event (Office)
keywords: vbaof11.chm299003
f1_keywords:
- vbaof11.chm299003
ms.prod: office
api_name:
- Office.CustomXMLParts.PartAfterLoad
ms.assetid: d59fe837-27b5-300f-133f-ffb01f5f95b9
ms.date: 06/08/2017
---


# CustomXMLParts.PartAfterLoad Event (Office)

Occurs just after a  **CustomXMLPart** object is loaded.


## Syntax

 _expression_. `PartAfterLoad`( `_Part_`, )

 _expression_ An expression that returns a [CustomXMLParts](./Office.CustomXMLParts.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Part_|Required|**CustomXMLPart**|The part that was loaded.|

## Example

The following example adds XML to a part after it is loaded.


```vb
Sub CustomXMLParts_PartAfterLoad(ByVal objPart As CustomXMLPart) 
   objPart.XML ("<root xmlns='https://www.w3c.org/XMLSchema'>text</root>") 
End Sub
```


## See also


[CustomXMLParts Object](Office.CustomXMLParts.md)



[CustomXMLParts Object Members](./overview/Library-Reference/customxmlparts-members-office.md)

