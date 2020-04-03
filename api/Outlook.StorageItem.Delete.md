---
title: StorageItem.Delete method (Outlook)
keywords: vbaol11.chm2145
f1_keywords:
- vbaol11.chm2145
ms.prod: outlook
api_name:
- Outlook.StorageItem.Delete
ms.assetid: 0ace6d9e-3dc7-52d5-ac20-97c2f3b109de
ms.date: 06/08/2017
localization_priority: Normal
---


# StorageItem.Delete method (Outlook)

Permanently removes the  **[StorageItem](Outlook.StorageItem.md)** object from the parent folder.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a [StorageItem](Outlook.StorageItem.md) object.


## Remarks

This call allows a solution to clean up or reset the storage for its private data. Attempting to delete a **StorageItem** that has been removed by a prior **StorageItem.Delete** call will result in the error, "Could not complete the deletion."

For more information on deleting solution data stored in a **StorageItem** object, see [Updating and Deleting Solution Storage](../outlook/How-to/Solution-Storage/updating-and-deleting-solution-storage.md).


## Example

The following code sample in Visual Basic for Applications shows how to clean up any existing  **StorageItem** object that has the specified subject, create a new instance with the same subject, assign a value to a custom property, and save the new instance.


```vb
Sub AssignStorageData() 
 
 Dim oInbox As Outlook.Folder 
 
 Dim myStorage As Outlook.StorageItem 
 
 
 
 Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 ' Remove and reset any existing instance of StorageItem of the specified subject 
 
 Set myStorage = oInbox.GetStorage("My Private Storage", olIdentifyBySubject) 
 
 myStorage.Delete 
 
 Set myStorage = Nothing 
 
 ' Get a new instance of StorageItem 
 
 Set myStorage = oInbox.GetStorage("My Private Storage", olIdentifyBySubject) 
 
 myStorage.UserProperties.Add "Order Number", olNumber 
 
 myStorage.UserProperties("Order Number").Value = 1000 
 
 myStorage.Save 
 
End Sub
```


## See also


[StorageItem Object](Outlook.StorageItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]