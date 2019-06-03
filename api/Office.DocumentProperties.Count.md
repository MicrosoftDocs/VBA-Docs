---
title: DocumentProperties.Count property (Office)
keywords: vbaof11.chm250013
f1_keywords:
- vbaof11.chm250013
ms.prod: office
api_name:
- Office.DocumentProperties.Count
ms.assetid: 8f4367bd-d30a-ba45-3ec2-3c5b94ede4d8
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperties.Count property (Office)

Gets a **Long** indicating the number of items in the **DocumentProperties** collection. Read-only.


## Syntax

_expression_.**Count** (_pc_)

_expression_ A variable that represents a **[DocumentProperties](Office.DocumentProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pc_|Required|**Long**|Represents the index of the document property.|

## Return value

Long


## Example

This example displays the number of custom document properties in the active document.


```vb
MsgBox ("There are " & _ 
 ActiveDocument.CustomDocumentProperties.Count & _ 
 " custom document properties in the " & _ 
 "active document.")
```


## See also

- [DocumentProperties object members](overview/library-reference/documentproperties-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]