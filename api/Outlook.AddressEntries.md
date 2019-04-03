---
title: AddressEntries object (Outlook)
keywords: vbaol11.chm24
f1_keywords:
- vbaol11.chm24
ms.prod: outlook
api_name:
- Outlook.AddressEntries
ms.assetid: db91b717-07c6-d1f2-c545-b766ee1f0c6b
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressEntries object (Outlook)

Contains a collection of addresses for an  **[AddressList](Outlook.AddressList.md)** object.


## Remarks

The object may contain zero or more  **[AddressEntry](Outlook.AddressEntry.md)** objects and provides access to the entries in a transport provider's address book container.


## Example

The following example sets a reference to an  **AddressEntries** object.






```vb
Set myNameSpace = Application.GetNameSpace("MAPI") 
 
Set myAddressList = myNameSpace.AddressLists("Personal Address Book") 
 
Set myAddressEntries = myAddressList.AddressEntries
```

You can also index directly into the  **AddressEntries** object, returning an **AddressEntry** object.




```vb
Set myAddressEntry = myAddressList.AddressEntries(index)
```


## Methods



|Name|
|:-----|
|[Add](Outlook.AddressEntries.Add.md)|
|[GetFirst](Outlook.AddressEntries.GetFirst.md)|
|[GetLast](Outlook.AddressEntries.GetLast.md)|
|[GetNext](Outlook.AddressEntries.GetNext.md)|
|[GetPrevious](Outlook.AddressEntries.GetPrevious.md)|
|[Item](Outlook.AddressEntries.Item.md)|
|[Sort](Outlook.AddressEntries.Sort.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.AddressEntries.Application.md)|
|[Class](Outlook.AddressEntries.Class.md)|
|[Count](Outlook.AddressEntries.Count.md)|
|[Parent](Outlook.AddressEntries.Parent.md)|
|[Session](Outlook.AddressEntries.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[AddressEntries Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
