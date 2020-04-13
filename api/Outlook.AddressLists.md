---
title: AddressLists object (Outlook)
keywords: vbaol11.chm87
f1_keywords:
- vbaol11.chm87
ms.prod: outlook
api_name:
- Outlook.AddressLists
ms.assetid: b8c5ce75-3030-0179-45bb-f44fe6628074
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressLists object (Outlook)

Contains a set of  **[AddressList](Outlook.AddressList.md)** objects.


## Remarks

The **AddressLists** collection provides access to the root of the transport provider's address book hierarchy for the current session.


## Example

The following example sets a reference to the  **AddressLists** object.


```vb
Set myAddressLists = myNameSpace.AddressLists
```


## Methods



|Name|
|:-----|
|[Item](Outlook.AddressLists.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.AddressLists.Application.md)|
|[Class](Outlook.AddressLists.Class.md)|
|[Count](Outlook.AddressLists.Count.md)|
|[Parent](Outlook.AddressLists.Parent.md)|
|[Session](Outlook.AddressLists.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]