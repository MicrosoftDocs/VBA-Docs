---
title: Add a Custom Folder to a Group and Display it in Overlay Mode by Default
ms.prod: outlook
ms.assetid: 79622092-bc9e-fd75-5579-dc626268d163
ms.date: 06/08/2017
localization_priority: Normal
---


# Add a Custom Folder to a Group and Display it in Overlay Mode by Default

You can add custom navigation folders to a navigation group in Microsoft Outlook by using the **[Add](../../../api/Outlook.NavigationFolders.Add.md)** method of the **[NavigationFolders](../../../api/Outlook.Folders.md)** collection for a **[NavigationGroup](../../../api/Outlook.NavigationGroup.md)** object. The **Add** method accepts a **[Folder](../../../api/Outlook.Folder.md)** object reference, to which the custom navigation folder is associated.

If the custom navigation folder is associated with a calendar folder, you can also use the **[IsSideBySide](../../../api/Outlook.NavigationFolder.IsSideBySide.md)** property of the **[NavigationFolder](../../../api/Outlook.Folder.md)** object to determine if the contents of the custom navigation folder are displayed in side-by-side or overlay mode.

This sample creates a new calendar folder for company events and adds a custom navigation folder for the new folder, configuring the custom navigation folder so that it is displayed by default in overlay mode.

The sample performs the following actions:

1. The sample obtains a **[Folder](../../../api/Outlook.Folder.md)** object reference for the **Calendar** default folder for the current user, by using the **[GetDefaultFolder](../../../api/Outlook.NameSpace.GetDefaultFolder.md)** method of the **[NameSpace](../../../api/Outlook.NameSpace.md)** object.
    
2. It then creates a new **Folder** object named "Company Events", representing the new calendar folder, in the **[Folders](../../../api/Outlook.Folders.md)** collection of the **Calendar** default folder.
    
3. The sample then obtains a reference to the **[NavigationPane](../../../api/Outlook.NavigationPane.md)** object for the active explorer and uses the **[GetNavigationModule](../../../api/Outlook.NavigationModules.GetNavigationModule.md)** method of the **[NavigationModules](../../../api/Outlook.NavigationModules.md)** collection to obtain a **[CalendarModule](../../../api/Outlook.CalendarModule.md)** object reference.
    
4. It then uses the **[GetDefaultNavigationGroup](../../../api/Outlook.NavigationGroups.GetDefaultNavigationGroup.md)** method of the **[NavigationGroups](../../../api/Outlook.NavigationGroups.md)** collection for the **CalendarModule** to obtains a **NavigationGroup** object reference to the **My Calendars** navigation group.
    
5. It then adds a new **NavigationFolder** object, based on the **Folder** object created earlier by the sample, to the navigation group by using the **Add** method of the **NavigationGroups** collection for that navigation group.
    
6. The sample then sets the **[CurrentModule](../../../api/Outlook.NavigationPane.CurrentModule.md)** property of the **NavigationPane** object to the **CalendarModule** object reference, to ensure that the **Calendar** navigation module is currently displayed in the Navigation Pane.
    
7. Finally, the sample then configures the navigation folder:
    
      - The sample sets the **[IsSelected](../../../api/Outlook.NavigationFolder.IsSelected.md)** property to **True** to display it in the active explorer.
    
  - The sample then sets the **IsSideBySide** property to **False** to display it by default in overlay mode.
    



```vb
Private Sub CreateCompanyEventsFolder() 
 Dim objNamespace As NameSpace 
 Dim objCalendar As Folder 
 Dim objFolder As Folder 
 
 Dim objPane As NavigationPane 
 Dim objModule As CalendarModule 
 Dim objGroup As NavigationGroup 
 Dim objNavFolder As NavigationFolder 
 
 On Error GoTo ErrRoutine 
 
 ' First, retrieve the default calendar folder. 
 Set objNamespace = Application.GetNamespace("MAPI") 
 Set objCalendar = objNamespace.GetDefaultFolder(olFolderCalendar) 
 
 ' Create a new calendar folder named "Company Events". 
 Set objFolder = objCalendar.Folders.Add("Company Events", olFolderCalendar) 
 
 ' Get the NavigationPane object for the 
 ' currently displayed Explorer object. 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
 ' Get the calendar module from the Navigation Pane. 
 Set objModule = objPane.Modules.GetNavigationModule(olModuleCalendar) 
 
 ' Get the "My Calendars" navigation group from the 
 ' calendar module. 
 With objModule.NavigationGroups 
 Set objGroup = .GetDefaultNavigationGroup(olMyFoldersGroup) 
 End With 
 
 ' Add a new navigation folder for the "Company Events" 
 ' folder in the "My Calendars" navigation group. 
 Set objNavFolder = objGroup.NavigationFolders.Add(objFolder) 
 
 ' Set the navigation folder to be displayed in overlay mode 
 ' by default. The IsSelected property can't be set to True 
 ' unless the CalendarModule object is the current module 
 ' displayed in the Navigation Pane. 
 Set objPane.CurrentModule = objModule 
 objNavFolder.IsSelected = True 
 objNavFolder.IsSideBySide = False 
 
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
 "CreateCompanyEventsFolder" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]