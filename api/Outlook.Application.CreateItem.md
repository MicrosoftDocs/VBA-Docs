---
title: Application.CreateItem Method (Outlook)
keywords: vbaol11.chm714
f1_keywords:
- vbaol11.chm714
ms.prod: outlook
api_name:
- Outlook.Application.CreateItem
ms.assetid: e5fbf367-db16-5042-823e-68e6b805e612
ms.date: 06/08/2017
---


# Application.CreateItem Method (Outlook)

Creates and returns a new Microsoft Outlook item.


## Syntax

 _expression_. `CreateItem`( `_ItemType_` )

 _expression_ A variable that represents an [Application](./Outlook.Application.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ItemType_|Required| **[OlItemType](Outlook.OlItemType.md)**|The Outlook item type for the new item.|

### Return Value

An  **Object** value that represents the new Outlook item.


## Remarks

The  **CreateItem** method can only create default Outlook items. To create new items using a custom form, use the **[Add](Outlook.Items.Add.md)** method on the **[Items](Outlook.Items.md)** collection.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new  **[MailItem](Outlook.MailItem.md)** object and sets the **BodyFormat** property to **olFormatHTML** . The Body text of the e-mail item will now appear in HTML format.


```vb
Sub CreateHTMLMail() 
 
 'Creates a new e-mail item and modifies its properties 
 
 Dim objMail As Outlook.MailItem 
 
 
 
 'Create e-mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 With objMail 
 
 'Set body format to HTML 
 
 .BodyFormat = olFormatHTML 
 
 .HTMLBody = "<HTML><H2>The body of this message will appear in HTML.</H2><BODY> Please enter the message text here. </BODY></HTML>" 
 
 .Display 
 
 End With 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)




[How to: Import Appointment XML Data into Outlook Appointment Objects](../outlook/How-to/Items-Folders-and-Stores/import-appointment-xml-data-into-outlook-appointment-objects-outlook.md)

