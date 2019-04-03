---
title: UserDefinedProperties object (Outlook)
keywords: vbaol11.chm3152
f1_keywords:
- vbaol11.chm3152
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperties
ms.assetid: 196e5d4c-22be-02d3-95e0-3ea7594c2e4b
ms.date: 06/08/2017
localization_priority: Normal
---


# UserDefinedProperties object (Outlook)

Contains a set of  **[UserDefinedProperty](Outlook.UserDefinedProperty.md)** objects representing the user-defined properties defined for a **[Folder](Outlook.Folder.md)** object.


## Remarks

The members of the  **UserDefinedProperties** collection correspond to the fields under **User-defined fields in folder** that you get in the **Show Fields** dialog.

Use the  **[UserDefinedProperties](Outlook.Folder.UserDefinedProperties.md)** property to retrieve the **UserDefinedProperties** object from a **Folder** object.

Use the  **[Add](Outlook.UserDefinedProperties.Add.md)** method to define and add a user-defined property to, and the **[Remove](Outlook.UserDefinedProperties.Remove.md)** method to remove an existing user-defined property from, the **UserDefinedProperties** collection. Use the **[Item](Outlook.UserDefinedProperties.Item.md)** method to retrieve by name or index, or the **[Find](Outlook.UserDefinedProperties.Find.md)** method to locate and retrieve by name, a **UserDefinedProperty** object from the **UserDefinedProperties** collection. Use the **[Refresh](Outlook.UserDefinedProperties.Refresh.md)** method to reload the **UserDefinedProperties** collection from the store.

The  **UserDefinedProperties** collection contains only the definitions of user-defined properties, which are applicable to all Outlook items contained by the folder. To retrieve or change user-defined property values for an Outlook item in that folder, use the **[UserProperties](Outlook.MailItem.UserProperties.md)** property of the Outlook item, such as a **[MailItem](Outlook.MailItem.md)** object, to retrieve the **[UserProperties](Outlook.UserProperties.md)** collection for that item. You can then use the **[UserProperty](Outlook.UserProperty.md)** object for the appropriate user-defined property to retrieve or change the value of that user-defined property for the Outlook item.


## Example

The following Visual Basic for Applications (VBA) example uses the  **Add** method to create and add several **UserDefinedProperty** objects to the **Inbox** default folder.


```vb
Sub AddStatusProperties() 
 
 Dim objNamespace As NameSpace 
 
 Dim objFolder As Folder 
 
 Dim objProperty As UserDefinedProperty 
 
 
 
 ' Obtain a Folder object reference to the 
 
 ' Inbox default folder. 
 
 Set objNamespace = Application.GetNamespace("MAPI") 
 
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox) 
 
 
 
 ' Add five user-defined properties, used to identify and 
 
 ' track customer issues. 
 
 With objFolder.UserDefinedProperties 
 
 Set objProperty = .Add("Issue?", olYesNo, olFormatYesNoIcon) 
 
 Set objProperty = .Add("Issue Research Time", olDuration) 
 
 Set objProperty = .Add("Issue Resolution Time", olDuration) 
 
 Set objProperty = .Add("Customer Follow-Up", olYesNo, olFormatYesNoYesNo) 
 
 Set objProperty = .Add("Issue Closed", olYesNo, olFormatYesNoYesNo) 
 
 End With 
 
End Sub 
 

```


## Methods



|Name|
|:-----|
|[Add](Outlook.UserDefinedProperties.Add.md)|
|[Find](Outlook.UserDefinedProperties.Find.md)|
|[Item](Outlook.UserDefinedProperties.Item.md)|
|[Refresh](Outlook.UserDefinedProperties.Refresh.md)|
|[Remove](Outlook.UserDefinedProperties.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.UserDefinedProperties.Application.md)|
|[Class](Outlook.UserDefinedProperties.Class.md)|
|[Count](Outlook.UserDefinedProperties.Count.md)|
|[Parent](Outlook.UserDefinedProperties.Parent.md)|
|[Session](Outlook.UserDefinedProperties.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]