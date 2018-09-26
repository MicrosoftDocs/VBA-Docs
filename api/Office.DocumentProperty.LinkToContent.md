---
title: DocumentProperty.LinkToContent Property (Office)
keywords: vbaof11.chm250008
f1_keywords:
- vbaof11.chm250008
ms.prod: office
api_name:
- Office.DocumentProperty.LinkToContent
ms.assetid: 062df6df-cdee-81fc-3244-e229dacaa64e
ms.date: 06/08/2017
---


# DocumentProperty.LinkToContent Property (Office)

Is  **True** if the value of the custom document property is linked to the content of the container document. **False** if the value is static. Read/write.


## Syntax

 _expression_. `LinkToContent`( `_pfLinkRetVal_` )

 _expression_ A variable that represents a [DocumentProperty](./Office.DocumentProperty.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pfLinkRetVal_|Required|**Boolean**|Indicates whether the document property is linked to the container document.|

## Remarks

This property applies only to custom document properties. For built-in document properties, the value of this property is  **False**.

Use the  **LinkSource** property to set the source for the specified linked property. Setting the **LinkSource** property sets the **LinkToContent** property to **True**.

For Excel, If LinkToContent is set to  **True**, you must supply an address or range name for the [LinkSource](./overview/Library-Reference.md) from the workbook. If the address or range name covers more than one cell, the custom document property takes the value from the top left cell of the range.


## Example

This example displays the linked status of the custom document property. For the example to work,  **dp** must be a valid **DocumentProperty** object.


```vb
Sub DisplayLinkStatus(dp As DocumentProperty) 
 Dim stat As String, tf As String 
 If dp.LinkToContent Then 
 tf = "" 
 Else 
 tf = "not " 
 End If 
 stat = "This property is " &amp; tf &amp; "linked" 
 If dp.LinkToContent Then 
 stat = stat + Chr(13) &amp; "The link source is " &amp; dp.LinkSource 
 End If 
 MsgBox stat 
End Sub
```


## See also


[DocumentProperty Object](Office.DocumentProperty.md)
[Sync Object](Office.Sync.md)



[DocumentProperty Object Members](./overview/Library-Reference/documentproperty-members-office.md)
[Sync Object Members](./overview/Library-Reference/sync-members-office.md)

