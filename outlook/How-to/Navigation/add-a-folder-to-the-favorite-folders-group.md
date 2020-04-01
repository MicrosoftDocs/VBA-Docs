---
title: Add a Folder to the Favorite Folders Group
ms.prod: outlook
ms.assetid: 5d0b448e-2f43-a58c-e44d-eecb9971f7ed
ms.date: 06/08/2017
localization_priority: Normal
---


# Add a Folder to the Favorite Folders Group

You can add a folder to the **Favorite Folders** navigation group in Microsoft Outlook by using the **[Add](../../../api/Outlook.NavigationFolders.Add.md)** method of the **[NavigationFolders](../../../api/Outlook.Folders.md)** collection for a **[NavigationGroup](../../../api/Outlook.NavigationGroup.md)** object. The **Add** method accepts a **[Folder](../../../api/Outlook.Folder.md)** object reference, to which the custom navigation folder is associated.

You can retrieve a **NavigationGroup** object reference to the default navigation group for a specified navigation group type by using the **[GetDefaultNavigationGroup](../../../api/Outlook.NavigationGroups.GetDefaultNavigationGroup.md)** method of the **NavigationGroups** object.

This sample creates a new mail folder for important items and adds a custom navigation folder for the new folder in the **Favorite Folders** navigation group of the **Mail** module.


 **Note** If you attempt to add a solution-specific folder, that is created for the Solutions module, to the Favorite Folders list, Outlook will raise an error.

The sample performs the following actions:

1. The sample obtains a **[Folder](../../../api/Outlook.Folder.md)** object reference for the **Inbox** default folder for the current user, by using the **[GetDefaultFolder](../../../api/Outlook.NameSpace.GetDefaultFolder.md)** method of the **[NameSpace](../../../api/Outlook.NameSpace.md)** object.
    
2. It then creates a new **Folder** object named "Important Items", representing the new mail folder, in the **[Folders](../../../api/Outlook.Folders.md)** collection of the **Inbox** default folder.
    
3. The sample then obtains a reference to the **[NavigationPane](../../../api/Outlook.NavigationPane.md)** object for the active explorer and uses the **[GetNavigationModule](../../../api/Outlook.NavigationModules.GetNavigationModule.md)** method of the **[NavigationModules](../../../api/Outlook.NavigationModules.md)** collection to obtain a **[MailModule](../../../api/Outlook.MailModule.md)** object reference.
    
4. It then uses the **[GetDefaultNavigationGroup](../../../api/Outlook.NavigationGroups.GetDefaultNavigationGroup.md)** method of the **[NavigationGroups](../../../api/Outlook.NavigationGroups.md)** collection for the **MailModule** to obtains a **NavigationGroup** object reference to the **Favorite Folders** navigation group.
    
5. Finally, the sample adds a new **NavigationFolder** object, based on the **Folder** object created earlier by the sample, to the navigation group by using the **Add** method of the **NavigationGroups** collection for that navigation group.
    



```vb
Private Sub CreateImportantFavoritesFolder() 
    Dim objNamespace As NameSpace 
    Dim objCalendars As Folder 
    Dim objFolder As Folder 
     
    Dim objPane As NavigationPane 
    Dim objModule As MailModule 
    Dim objGroup As NavigationGroup 
    Dim objNavFolder As NavigationFolder 
     
    On Error GoTo ErrRoutine 
     
    ' First, retrieve the default Inbox folder. 
    Set objNamespace = Application.GetNamespace("MAPI") 
    Set objCalendars = objNamespace.GetDefaultFolder(olFolderInbox) 
     
    ' Create a new mail folder named "Important Items". 
    Set objFolder = objCalendars.Folders.Add("Important Items") 
         
    ' Get the NavigationPane object for the 
    ' currently displayed Explorer object. 
    Set objPane = Application.ActiveExplorer.NavigationPane 
     
    ' Get the mail module from the Navigation Pane. 
    Set objModule = objPane.Modules.GetNavigationModule(olModuleMail) 
     
    ' Get the "Favorite Folders" navigation group from the 
    ' mail module. 
    With objModule.NavigationGroups 
        Set objGroup = .GetDefaultNavigationGroup(olFavoriteFoldersGroup) 
    End With 
     
    ' Add a new navigation folder for the "Important Items" 
    ' folder in the "Favorite Folders" navigation group. 
    Set objNavFolder = objGroup.NavigationFolders.Add(objFolder) 
     
EndRoutine: 
    On Error GoTo 0 
    Set objNavFolder = Nothing 
    Set objFolder = Nothing 
    Set objGroup = Nothing 
    Set objModule = Nothing 
    Set objPane = Nothing 
    Set objNamespace = Nothing 
    Exit Sub 
 
ErrRoutine: 
    MsgBox Err.Number & " - " & Err.Description, _ 
        vbOKOnly Or vbCritical, _ 
        "CreateImportantFavoritesFolder" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]