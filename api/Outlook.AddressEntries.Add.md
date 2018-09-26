---
title: AddressEntries.Add Method (Outlook)
keywords: vbaol11.chm32
f1_keywords:
- vbaol11.chm32
ms.prod: outlook
api_name:
- Outlook.AddressEntries.Add
ms.assetid: b4c37547-8fbd-b1e4-40f3-5cba3cffd6e9
ms.date: 06/08/2017
---


# AddressEntries.Add Method (Outlook)

Adds a new entry to the  **[AddressEntries](Outlook.AddressEntries.md)** collection.


## Syntax

 _expression_. `Add`( `_Type_` , `_Name_` , `_Address_` )

 _expression_ An [AddressEntries](./Outlook.AddressEntries.md) object that represents the new entry.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **String**|The type of the new entry.|
| _Name_|Optional| **Variant**|The name of the new entry.|
| _Address_|Optional| **Variant**|The address.|

### Return value

An  **[AddressEntry](Outlook.AddressEntry.md)** object that represents the new entry.


## Remarks

New entries or changes to existing entries are not persisted in the collection until after calling the  **[Update](Outlook.AddressEntry.Update.md)** method.


## See also


[AddressEntries Object](Outlook.AddressEntries.md)

