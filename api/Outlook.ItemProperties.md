---
title: ItemProperties object (Outlook)
keywords: vbaol11.chm530
f1_keywords:
- vbaol11.chm530
ms.prod: outlook
api_name:
- Outlook.ItemProperties
ms.assetid: 34a110ed-6617-72da-1e98-a9773c705b40
ms.date: 06/08/2017
localization_priority: Normal
---


# ItemProperties object (Outlook)

A collection of all properties associated with the item.


## Remarks

Use the **[ItemProperties](Outlook.MailItem.ItemProperties.md)** property to return the **ItemProperties** collection. Use **ItemProperties.Item** (_index_), where _index_ is the name of the object or the numeric position of the item within the collection, to return a single **[ItemProperty](Outlook.ItemProperty.md)** object.


> [!NOTE] 
> The **ItemProperties** collection is zero-based, meaning that the first item in the collection is referenced by 0.

Use the **[Add](Outlook.ItemProperties.Add.md)** method to add a new item property to the **ItemProperties** collection. Use the **[Remove](Outlook.ItemProperties.Remove.md)** method to remove an item property from the **ItemProperties** collection.


> [!NOTE] 
>  You can only add or remove custom properties. Custom properties are denoted by the **[IsUserProperty](Outlook.ItemProperty.IsUserProperty.md)**.


## Example

The following example creates a new **[MailItem](Outlook.MailItem.md)** object and stores its **ItemProperties** collection in a variable called `objItems`.


```vb
Sub ItemProperty() 
 
 'Creates a new MailItem and access its properties 
 
 Dim objMail As MailItem 
 
 Dim objItems As ItemProperties 
 
 Dim objItem As ItemProperty 
 
 
 
 'Create the mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 'Create a reference to the item properties collection 
 
 Set objItems = objMail.ItemProperties 
 
 'Create a reference to the item property page 
 
 Set objItem = objItems.item(0) 
 
End Sub
```


## Methods



|Name|
|:-----|
|[Add](Outlook.ItemProperties.Add.md)|
|[Item](Outlook.ItemProperties.Item.md)|
|[Remove](Outlook.ItemProperties.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.ItemProperties.Application.md)|
|[Class](Outlook.ItemProperties.Class.md)|
|[Count](Outlook.ItemProperties.Count.md)|
|[Parent](Outlook.ItemProperties.Parent.md)|
|[Session](Outlook.ItemProperties.Session.md)|

## See also


[ItemProperties Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]