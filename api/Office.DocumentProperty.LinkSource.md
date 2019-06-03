---
title: DocumentProperty.LinkSource property (Office)
keywords: vbaof11.chm250009
f1_keywords:
- vbaof11.chm250009
ms.prod: office
api_name:
- Office.DocumentProperty.LinkSource
ms.assetid: 3e3a6ebc-615a-298e-c40f-cbb6d5cf63e3
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperty.LinkSource property (Office)

Gets or sets the source of a linked custom document property. Read/write.


## Syntax

_expression_.**LinkSource**(_pbstrSourceRetVal_)

_expression_ A variable that represents a **[DocumentProperty](Office.DocumentProperty.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pbstrSourceRetVal_|Required|**String**|Represents the name of the source of the document property.|

## Remarks

This property applies only to custom document properties; you cannot use it with built-in document properties.

The source of the specified link is defined by the container application.

Setting the **LinkSource** property sets the **LinkToContent** property to **True**.


## Example

This example displays the linked status of a custom document property. For the example to work, **dp** must be a valid **DocumentProperty** object.


```vb
Sub DisplayLinkStatus(dp As DocumentProperty) 
 Dim stat As String, tf As String 
 If dp.LinkToContent Then 
 tf = "" 
 Else 
 tf = "not " 
 End If 
 stat = "This property is " & tf & "linked" 
 If dp.LinkToContent Then 
 stat = stat + Chr(13) & "The link source is " & dp.LinkSource 
 End If 
 MsgBox stat 
End Sub
```


## See also

- [DocumentProperty object members](overview/library-reference/documentproperty-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]