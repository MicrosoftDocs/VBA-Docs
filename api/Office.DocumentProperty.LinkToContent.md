---
title: DocumentProperty.LinkToContent property (Office)
keywords: vbaof11.chm250008
f1_keywords:
- vbaof11.chm250008
ms.prod: office
api_name:
- Office.DocumentProperty.LinkToContent
ms.assetid: 062df6df-cdee-81fc-3244-e229dacaa64e
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperty.LinkToContent property (Office)

Is **True** if the value of the custom document property is linked to the content of the container document. **False** if the value is static. Read/write.


## Syntax

_expression_.**LinkToContent**(_pfLinkRetVal_)

_expression_ A variable that represents a **[DocumentProperty](Office.DocumentProperty.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pfLinkRetVal_|Required|**Boolean**|Indicates whether the document property is linked to the container document.|

## Remarks

This property applies only to custom document properties. For built-in document properties, the value of this property is **False**.

Use the **[LinkSource](office.documentproperty.linksource.md)** property to set the source for the specified linked property. Setting the **LinkSource** property sets the **LinkToContent** property to **True**.

For Excel, if **LinkToContent** is set to **True**, you must supply an address or range name for the **LinkSource** from the workbook. If the address or range name covers more than one cell, the custom document property takes the value from the top left cell of the range.


## Example

This example displays the linked status of the custom document property. For the example to work, **dp** must be a valid **DocumentProperty** object.


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
- [Sync object](Office.Sync.md)
- [Sync object members](overview/Library-Reference/sync-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]