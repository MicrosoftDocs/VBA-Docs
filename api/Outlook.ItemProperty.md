---
title: ItemProperty object (Outlook)
keywords: vbaol11.chm517
f1_keywords:
- vbaol11.chm517
ms.prod: outlook
api_name:
- Outlook.ItemProperty
ms.assetid: 3570d1f9-40ed-0a99-f63c-141134418c3b
ms.date: 06/08/2017
localization_priority: Normal
---


# ItemProperty object (Outlook)

Represents information about a given item property for a Microsoft Outlook item object.


## Remarks

 Each item property defines a certain attribute of the item, such as the name, type, or value of the item. The **ItemProperty** object is a member of the **[ItemProperties](Outlook.ItemProperties.md)** collection.

Use  **ItemProperties.Item** (_index_), where _index_ is the object's numeric position within the collection or it's name to return a single **ItemProperty** object.


## Example

The following example creates a reference to the first  **ItemProperty** object in the **ItemProperties** collection.


```vb
Sub NewMail() 
 
 'Creates a new MailItem and references the ItemProperties collection. 
 
 Dim objMail As MailItem 
 
 Dim objitems As ItemProperties 
 
 Dim objitem As ItemProperty 
 
 
 
 'Create a new mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 'Create a reference to the ItemProperties collection 
 
 Set objitems = objMail.ItemProperties 
 
 'Create reference to the first object in the collection 
 
 Set objitem = objitems.item(0) 
 
End Sub
```


## Properties



|Name|
|:-----|
|[Application](Outlook.ItemProperty.Application.md)|
|[Class](Outlook.ItemProperty.Class.md)|
|[IsUserProperty](Outlook.ItemProperty.IsUserProperty.md)|
|[Name](Outlook.ItemProperty.Name.md)|
|[Parent](Outlook.ItemProperty.Parent.md)|
|[Session](Outlook.ItemProperty.Session.md)|
|[Type](Outlook.ItemProperty.Type.md)|
|[Value](Outlook.ItemProperty.Value.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[ItemProperty Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]