---
title: Delete Data to Reset Solution Storage
ms.prod: outlook
ms.assetid: 38147c59-3145-3df1-7488-1df26ba0e1fa
ms.date: 06/08/2019
localization_priority: Normal
---


# Delete Data to Reset Solution Storage

This topic describes how to delete existing solution data to reset the solution storage:


1. Use **[Folder.GetStorage](../../../api/Outlook.Folder.GetStorage.md)** to obtain an existing **[StorageItem](../../../api/Outlook.StorageItem.md)** object in a specific folder. This call will return a new **StorageItem** object if none already exists.

2. Use **[StorageItem.Delete](../../../api/Outlook.StorageItem.Delete.md)** to remove the object permanently from the folder.

3. Use **Folder.GetStorage** to create a new instance of the **StorageItem** object with the same subject.

4. Use the **[Add](../../../api/Outlook.UserProperties.Add.md)** method of **[StorageItem.UserProperties](../../../api/Outlook.StorageItem.UserProperties.md)** to create a custom property **Order Number**.

5. Set the **Order Number** property.

6. Use **[StorageItem.Save](../../../api/Outlook.StorageItem.Save.md)** to save the **StorageItem** object to the folder.


```vb
Sub StoreData() 
 Dim oInbox As Folder 
 Dim myStorage As StorageItem 
 Dim myPrivateProperty As UserProperty 
 
 Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 ' Get an existing instance of StorageItem by subject 
 Set myStorage = oInbox.GetStorage("My Private Storage", olIdentifyBySubject) 
 
 'Remove the storage permanently assuming it's old 
 myStorage.Delete 
 Set myStorage = Nothing 
 
 'Get a new instance of StorageItem in the Inbox 
 Set myStorage = oInbox.GetStorage("My Private Storage", olIdentifyBySubject) 
 
 'Create custom property for Order Number 
 Set myPrivateProperty = myStorage.UserProperties.Add("Order Number", olNumber) 
 
 'Store application data in the Order Number property 
 myPrivateProperty.Value = 1000 
 
 'Save the data to the Inbox 
 myStorage.Save 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]