---
title: DocumentProperties.Count property (Office)
keywords: vbaof11.chm250013
f1_keywords:
- vbaof11.chm250013
ms.prod: office
api_name:
- Office.DocumentProperties.Count
ms.assetid: 8f4367bd-d30a-ba45-3ec2-3c5b94ede4d8
ms.date: 06/08/2017
---


# DocumentProperties.Count property (Office)

Gets a  **Long** indicating the number of items in the **DocumentProperties** collection. Read-only.


## Syntax

_expression_. `Count`( `_pc_` )

_expression_ A variable that represents a [DocumentProperties](Office.DocumentProperties.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pc_|Required|**Long**|Represents the index of the document property.|

## Return value

Long


## Example

This example displays the number of custom document properties in the active document.


```vb
MsgBox ("There are " &amp; _ 
 ActiveDocument.CustomDocumentProperties.Count &amp; _ 
 " custom document properties in the " &amp; _ 
 "active document.")
```


## See also


[DocumentProperties Object](Office.DocumentProperties.md)



[DocumentProperties Object Members](./overview/Library-Reference/documentproperties-members-office.md)

