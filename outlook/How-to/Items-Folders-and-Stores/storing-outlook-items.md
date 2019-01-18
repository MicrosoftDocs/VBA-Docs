---
title: Storing Outlook items
ms.prod: outlook
ms.assetid: e4a639a4-10b2-7665-9261-19d6e7707e48
ms.date: 06/08/2017
localization_priority: Normal
---


# Storing Outlook Items

This topic describes how Outlook items are stored in folders and stores based on an account in the current profile.

The Outlook object model provides the following objects to store Outlook items:

- The  **[Folder](../../../api/Outlook.Folder.md)** object, which represents a container for other **Folder** objects and Outlook items.
    
     **Note**  The  **Folder** object has replaced the **MAPIFolder** object that existed in Microsoft Office Outlook 2003 and earlier versions of Outlook. New solutions should only use **Folder**.
- The  **[Folders](../../../api/Outlook.Folders.md)** collection, which represents all the **Folder** objects at one level of the folder tree in a store. The **Folders** collection can also represent a collection of search folders.
    
     **Note**  Although a search folder is represented programmatically by a  **Folder** object, not all events, methods, and properties of **Folder** apply to search folders.
- The  **[Store](../../../api/Outlook.Store.md)** object, which represents a file on the local computer or a network drive that stores email messages and other items. If you use an Exchange server, you can have a store on the server, in an Exchange Public folder, or on a local computer in a Personal Folders File (.pst) or Offline Folder File (.ost). For a POP3, IMAP, and HTTP email server, a store is a .pst file.
    
    You can add a store to the current profile using  **[NameSpace.AddStore](../../../api/Outlook.NameSpace.AddStore.md)** and **[NameSpace.AddStoreEx](../../../api/Outlook.NameSpace.AddStoreEx.md)**, and remove an existing store from the current profile using  **[NameSpace.RemoveStore](../../../api/Outlook.NameSpace.RemoveStore.md)**.
    
- The  **[Stores](../../../api/Outlook.Stores.md)** collection, which represents all the stores in the current Outlook profile. A profile defines one or more email accounts, and each email account is associated with a server of a specific type. The type of server determines the type of the store and how email and other items are delivered and stored. For example, an Exchange server stores email and other items in either a .pst file or a .ost file on the local computer or a mapped network drive, and an HTTP server (such as Hotmail) stores items in a .pst file on the local computer.
    

The  **Store** and **Stores** objects support the following:

- Enumerating folders in a store using  **[Store.GetRootFolder](../../../api/Outlook.Store.GetRootFolder.md)** and then **[Folder.Folders](../../../api/Outlook.Folder.Folders.md)**.
    
- Enumerating search folders in a store using  **[Store.GetSearchFolders](../../../api/Outlook.Store.GetSearchFolders.md)**.
    
     **Note**  Since a store does not necessarily support search folders, in general, you should trap for returned errors when using  **Store.GetSearchFolders** to obtain any search folders on a store.
- Better performance with enumerating folders. Because getting the root folder or search folders in a store requires the store to be open and opening a store imposes an overhead on performance, you can check the  **[Store.IsOpen](../../../api/Outlook.Store.IsOpen.md)** property before you decide to pursue the operation.
    
- Locating a local store (.pst or .ost) for an Exchange server, or a store (.pst) for a POP3, IMAP, or HTTP email server, using the  **[Store.FilePath](../../../api/Outlook.Store.FilePath.md)** property.
    
- Discovery of the Exchange store type and differentiation among different Exchange store types using the  **[Store.ExchangeStoreType](../../../api/Outlook.Store.ExchangeStoreType.md)** property.
    
- Additional information for an Exchange server through the  **[Store.IsCachedExchange](../../../api/Outlook.Store.IsCachedExchange.md)** and **[Store.IsDataFileStore](../../../api/Outlook.Store.IsDataFileStore.md)** properties.
    
- The  **[PropertyAccessor](../../../api/Outlook.PropertyAccessor.md)** object through the **[Store.PropertyAccessor](../../../api/Outlook.Store.PropertyAccessor.md)** property, allowing access to store properties that are not exposed as explicit built-in properties in the Outlook object model.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]