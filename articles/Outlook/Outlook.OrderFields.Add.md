---
title: OrderFields.Add Method (Outlook)
keywords: vbaol11.chm2678
f1_keywords:
- vbaol11.chm2678
ms.prod: outlook
api_name:
- Outlook.OrderFields.Add
ms.assetid: aabd32ef-e707-ddc5-24b6-723293273e56
ms.date: 06/08/2017
---


# OrderFields.Add Method (Outlook)

Creates a new  **[OrderField](Outlook.OrderField.md)** object and appends it to the **[OrderFields](Outlook.OrderFields.md)** collection.


## Syntax

 _expression_ . **Add**( **_PropertyName_** , **_IsDescending_** )

 _expression_ A variable that represents an **OrderFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PropertyName_|Required| **String**|The name of the property to which the new object is associated.|
| _IsDescending_|Optional| **Boolean**|The value used to set the  **[IsDescending](Outlook.OrderField.IsDescending.md)** property of the new **OrderField** object. If this value is not specified, the default value of the **IsDescending** property is used.|

### Return Value

An  **OrderField** object that represents the new order field.


## See also


#### Concepts


[OrderFields Object](Outlook.OrderFields.md)

