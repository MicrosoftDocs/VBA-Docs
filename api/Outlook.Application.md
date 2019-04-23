---
title: Application object (Outlook)
keywords: vbaol11.chm2991
f1_keywords:
- vbaol11.chm2991
ms.prod: outlook
api_name:
- Outlook.Application
ms.assetid: 797003e7-ecd1-eccb-eaaf-32d6ddde8348
ms.date: 06/08/2017
localization_priority: Normal
---


# Application object (Outlook)

Represents the entire Microsoft Outlook application.


## Remarks

 This is the only object in the hierarchy that can be returned by using the **[CreateObject](Outlook.Application.CreateObject.md)** method or the intrinsic Visual Basic **GetObject** function.

The Outlook  **Application** object has several purposes:


- As the root object, it allows access to other objects in the Outlook hierarchy.
    
- It allows direct access to a new item created by using  **[CreateItem](Outlook.Application.CreateItem.md)**, without having to traverse the object hierarchy.
    
- It allows access to the active interface objects (the explorer and the inspector).
    
When you use Automation to control Outlook from another application, you use the  **CreateObject** method to create an Outlook **Application** object.


## Example

The following Visual Basic for Applications (VBA) example starts Outlook (if it's not already running) and opens the default Inbox folder.


```vb
Set myNameSpace = Application.GetNameSpace("MAPI") 
 
Set myFolder= _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
myFolder.Display
```

The following Visual Basic for Applications (VBA) example uses the  **Application** object to create and open a new contact.




```vb
Set myItem = Application.CreateItem(olContactItem) 
 
myItem.Display
```


## Events



|Name|
|:-----|
|[AdvancedSearchComplete](Outlook.Application.AdvancedSearchComplete.md)|
|[AdvancedSearchStopped](Outlook.Application.AdvancedSearchStopped.md)|
|[BeforeFolderSharingDialog](Outlook.Application.BeforeFolderSharingDialog.md)|
|[ItemLoad](Outlook.Application.ItemLoad.md)|
|[ItemSend](Outlook.Application.ItemSend.md)|
|[MAPILogonComplete](Outlook.Application.MAPILogonComplete.md)|
|[NewMail](Outlook.Application.NewMail.md)|
|[NewMailEx](Outlook.Application.NewMailEx.md)|
|[OptionsPagesAdd](Outlook.Application.OptionsPagesAdd.md)|
|[Quit](Outlook.Application.Quit(even).md)|
|[Reminder](Outlook.Application.Reminder.md)|
|[Startup](Outlook.Application.Startup.md)|

## Methods



|Name|
|:-----|
|[ActiveExplorer](Outlook.Application.ActiveExplorer.md)|
|[ActiveInspector](Outlook.Application.ActiveInspector.md)|
|[ActiveWindow](Outlook.Application.ActiveWindow.md)|
|[AdvancedSearch](Outlook.Application.AdvancedSearch.md)|
|[CopyFile](Outlook.Application.CopyFile.md)|
|[CreateItem](Outlook.Application.CreateItem.md)|
|[CreateItemFromTemplate](Outlook.Application.CreateItemFromTemplate.md)|
|[CreateObject](Outlook.Application.CreateObject.md)|
|[GetNamespace](Outlook.Application.GetNamespace.md)|
|[GetObjectReference](Outlook.Application.GetObjectReference.md)|
|[IsSearchSynchronous](Outlook.Application.IsSearchSynchronous.md)|
|[Quit](Outlook.Application.Quit(method).md)|
|[RefreshFormRegionDefinition](Outlook.Application.RefreshFormRegionDefinition.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Application.Application.md)|
|[Assistance](Outlook.Application.Assistance.md)|
|[Class](Outlook.Application.Class.md)|
|[COMAddIns](Outlook.Application.COMAddIns.md)|
|[DefaultProfileName](Outlook.Application.DefaultProfileName.md)|
|[Explorers](Outlook.Application.Explorers.md)|
|[Inspectors](Outlook.Application.Inspectors.md)|
|[IsTrusted](Outlook.Application.IsTrusted.md)|
|[LanguageSettings](Outlook.Application.LanguageSettings.md)|
|[Name](Outlook.Application.Name.md)|
|[Parent](Outlook.Application.Parent.md)|
|[PickerDialog](Outlook.Application.PickerDialog.md)|
|[ProductCode](Outlook.Application.ProductCode.md)|
|[Reminders](Outlook.Application.Reminders.md)|
|[Session](Outlook.Application.Session.md)|
|[TimeZones](Outlook.Application.TimeZones.md)|
|[Version](Outlook.Application.Version.md)|

## See also

[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
