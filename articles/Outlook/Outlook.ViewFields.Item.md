---
title: ViewFields.Item Method (Outlook)
keywords: vbaol11.chm2551
f1_keywords:
- vbaol11.chm2551
ms.prod: outlook
api_name:
- Outlook.ViewFields.Item
ms.assetid: 5b7072b7-5f5e-2a39-1001-0b103a287a78
ms.date: 06/08/2017
---


# ViewFields.Item Method (Outlook)

Returns a  **[ViewField](Outlook.ViewField.md)** object from the collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **ViewFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The value can be a one-based integer that indexes an  **ViewField** object in the **[ViewFields](Outlook.ViewFields.md)** collection, a string that matches the **[ViewXMLSchemaName](Outlook.ViewField.ViewXMLSchemaName.md)** property value of an **ViewField** object in the collection, or a field name as displayed in the **Field Chooser**.|

### Return Value

A  **ViewField** object that represents the specified object.


## See also


#### Concepts


[ViewFields Object](Outlook.ViewFields.md)

