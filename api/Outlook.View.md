---
title: View object (Outlook)
keywords: vbaol11.chm2479
f1_keywords:
- vbaol11.chm2479
ms.prod: outlook
api_name:
- Outlook.View
ms.assetid: 41c8d149-9912-1685-4c8b-3c849cc6f1ed
ms.date: 06/08/2017
localization_priority: Normal
---


# View object (Outlook)

Represents a customizable view used to sort, group, and view data.


## Remarks

The  **View** object allows you to create customizable views that allow you to better sort, group and ultimately view data of all different types. There are a variety of different view types that provide the flexibility needed to create and maintain your important data.


- The table view type (**olTableView**) allows you to view data in a simple field-based table.
    
- The Calendar view type (**olCalendarView**) allows you to view data in a calendar format.
    
- The card view type (**olCardView**) allows you to view data in a series of cards. Each card displays the information contained by the item and can be sorted.
    
- The icon view type (**olIconView**) allows you to view data as icons, similar to a Windows folder or explorer.
    
- The timeline view type (**olTimelineView**) allows you to view data as it is received in a customizable linear time line.
    
Views are defined and customized using the  **View** object's **[XML](Outlook.View.XML.md)** property. The **XML** property allows you to create and set a customized XML schema that defines the various features of a view.

Use  **Views** (_index_), where _index_ is the name of the **View** object or its ordinal value, to return a single **View** object.

Use the  **[Add](Outlook.Views.Add.md)** method of the **Views** collection to create a new view.

Always use  **[Save](Outlook.View.Save.md)** to save a view after you change any property of the view.


## Example

The following example returns a view called Table View and stores it in a variable of type  **View** called objView. Before running this example, make sure a view by the name 'Table View' exists.


```vb
Sub GetView() 
 
 'Creates a new view 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 'Return a view called Table View 
 
 Set objView = objViews.Item("Table View") 
 
End Sub
```

The following example creates a new view of type  **olTableView** called New Table.




```vb
Sub CreateView() 
 
 'Creates a new view 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objNewView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 Set objNewView = objViews.Add(Name:="New Table", _ 
 
 ViewType:=olTableView, SaveOption:=olViewSaveOptionThisFolderEveryone) 
 
End Sub
```


## Methods



|Name|
|:-----|
|[Apply](Outlook.View.Apply.md)|
|[Copy](Outlook.View.Copy.md)|
|[Delete](Outlook.View.Delete.md)|
|[GoToDate](Outlook.View.GoToDate.md)|
|[Reset](Outlook.View.Reset.md)|
|[Save](Outlook.View.Save.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.View.Application.md)|
|[Class](Outlook.View.Class.md)|
|[Filter](Outlook.View.Filter.md)|
|[Language](Outlook.View.Language.md)|
|[LockUserChanges](Outlook.View.LockUserChanges.md)|
|[Name](Outlook.View.Name.md)|
|[Parent](Outlook.View.Parent.md)|
|[SaveOption](Outlook.View.SaveOption.md)|
|[Session](Outlook.View.Session.md)|
|[Standard](Outlook.View.Standard.md)|
|[ViewType](Outlook.View.ViewType.md)|
|[XML](Outlook.View.XML.md)|

## See also


[View Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]