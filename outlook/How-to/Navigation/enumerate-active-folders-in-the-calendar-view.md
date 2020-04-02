---
title: Enumerate Active Folders in the Calendar View
ms.prod: outlook
ms.assetid: 379bd7c7-d0bc-856f-4432-17e38342611b
ms.date: 06/08/2019
localization_priority: Normal
---


# Enumerate Active Folders in the Calendar View

In Microsoft Outlook, you can traverse the group and folder hierarchy of a module in the Navigation Pane by using the **[NavigationGroups](../../../api/Outlook.NavigationGroups.md)** and **[NavigationFolders](../../../api/Outlook.NavigationFolders.md)** collections. The **NavigationGroups** collection of the **[NavigationModule](../../../api/Outlook.NavigationModule.md)** object contains each navigation group displayed in a navigation module, while the **NavigationFolders** collection of the **[NavigationGroup](../../../api/Outlook.NavigationGroup.md)** object contains each navigation folder displayed in a navigation group.

By using these collections in combination, you can enumerate each navigation folder for a navigation module displayed in the Navigation Pane. 

The following sample counts the number of navigation folders selected for display in the **Calendar** navigation module of the Navigation Pane. The sample performs the following actions:


1. The sample first obtains a reference to the **[NavigationPane](../../../api/Outlook.NavigationPane.md)** object for the active explorer.
    
2. It then uses the **[GetNavigationModule](../../../api/Outlook.NavigationModules.GetNavigationModule.md)** method of the **[NavigationModules](../../../api/Outlook.NavigationModules.md)** collection to obtain a reference to the **[CalendarModule](../../../api/Outlook.CalendarModule.md)** object from the **NavigationPane** object.
    
3. The sample then enumerates through the **[NavigationGroups](../../../api/Outlook.CalendarModule.NavigationGroups.md)** collection of the **CalendarModule** object. For each **NavigationGroup** in the collection, the sample then enumerates the **[NavigationFolders](../../../api/Outlook.NavigationGroup.NavigationFolders.md)** collection.
    
4. If the **[IsSelected](../../../api/Outlook.NavigationFolder.IsSelected.md)** property of a **NavigationFolder** object contained in the **NavigationFolders** collection is set to **True**, the variable  `intCounter` is incremented.
    
5. Finally, the sample displays a dialog box containing the value of  `intCounter`.
    



```vb
Dim WithEvents objPane As NavigationPane 
 
Private Sub EnumerateActiveCalendarFolders() 
 Dim objModule As CalendarModule 
 Dim objGroup As NavigationGroup 
 Dim objFolder As NavigationFolder 
 Dim intCounter As Integer 
 
 On Error GoTo ErrRoutine 
 
 ' Get the NavigationPane object for the 
 ' currently displayed Explorer object. 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
 ' Get the CalendarModule object, if one exists, 
 ' for the current Navigation Pane. 
 Set objModule = objPane.Modules.GetNavigationModule(olModuleCalendar) 
 
 ' Iterate through each NavigationGroup contained 
 ' by the CalendarModule. 
 For Each objGroup In objModule.NavigationGroups 
 ' Iterate through each NavigationFolder contained 
 ' by the NavigationGroup. 
 For Each objFolder In objGroup.NavigationFolders 
 ' Check if the folder is selected. 
 If objFolder.IsSelected Then 
 intCounter = intCounter + 1 
 End If 
 Next 
 Next 
 
 ' Display the results. 
 MsgBox "There are " & intCounter & " selected calendars in the Calendar module." 
 
EndRoutine: 
 On Error GoTo 0 
 Set objFolder = Nothing 
 Set objGroup = Nothing 
 Set objModule = Nothing 
 Set objPane = Nothing 
 intCounter = 0 
 Exit Sub 
 
ErrRoutine: 
 MsgBox Err.Number & " - " & Err.Description, _ 
 vbOKOnly Or vbCritical, _ 
 "EnumerateActiveCalendarFolders" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]