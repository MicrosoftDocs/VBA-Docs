---
title: Store Data in a StorageItem for a Solution
ms.prod: outlook
ms.assetid: 75adfdbe-1c4d-fbd0-22ea-8f8fd5e212a5
ms.date: 06/08/2017
localization_priority: Normal
---


# Store Data in a StorageItem for a Solution

This topic describes how to store private application data in solution storage provided by the Outlook object model.


1. Determine the folder where you would like to store your application data. 
    
     **Note**  Because solution storage is created as hidden items in a folder, you can only store solution data if the store provider supports hidden items and the client has rights to write to that folder.
2. Use  **[Folder.GetStorage](../../../api/Outlook.Folder.GetStorage.md)** to obtain either an existing **[StorageItem](../../../api/Outlook.StorageItem.md)** object or a new **StorageItem** object if one does not already exist.
    
3. Use  **[StorageItem.Size](../../../api/Outlook.StorageItem.Size.md)** to determine if the **StorageItem** is new. If it is, then use the **[Add](../../../api/Outlook.UserProperties.Add.md)** method of **[StorageItem.UserProperties](../../../api/Outlook.StorageItem.UserProperties.md)** to create a custom property **Order Number**.
    
4. Set the  **Order Number** property. This assumes that an existing **StorageItem** already has the custom property **Order Number** defined.
    
5. Use  **[StorageItem.Save](../../../api/Outlook.StorageItem.Save.md)** to save the **StorageItem** object as a hidden item in the folder.
    

```vb
Sub StoreData() 
 Dim oInbox As Folder 
 Dim myStorage As StorageItem 
 Dim myPrivateProperty As UserProperty 
 
 Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 ' Get an existing instance of StorageItem by subject, or create new if it doesn't exist 
 Set myStorage = oInbox.GetStorage("My Private Storage", olIdentifyBySubject) 
 
 If myStorage.Size = 0 Then 
 'There was no existing StorageItem by this subject, so created a new one 
 'Create a custom property for Order Number 
 Set myPrivateProperty = myStorage.UserProperties.Add("Order Number", olNumber) 
 Else 
 'Assume that existing storage has the Order Number property already 
 Set myPrivateProperty = myStorage.UserProperties("Order Number") 
 End If 
 myPrivateProperty.Value = lngOrderNumber 
 myStorage.Save 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]