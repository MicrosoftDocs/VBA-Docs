---
title: Views object (Outlook)
keywords: vbaol11.chm3013
f1_keywords:
- vbaol11.chm3013
ms.prod: outlook
api_name:
- Outlook.Views
ms.assetid: 5dd7edc2-12a2-f4c2-d158-8053d80e8dc9
ms.date: 06/08/2017
localization_priority: Normal
---


# Views object (Outlook)

Contains a collection of all  **[View](Outlook.View.md)** objects in the current folder.


## Remarks

Use the  **Views** property of the **[Folder](Outlook.Folder.md)** object to return the **Views** collection. Use **Views** (_index_),where _index_ is the object's name or position within the collection, to return a single **View** object.

Use the  **[Add](Outlook.Views.Add.md)** method of the views collection to add a new view to the collection.

Use the  **[Remove](Outlook.Views.Remove.md)** method to remove a view from the collection.


## Example

The following example returns a  **View** object of type **olTableView** called Table View. Before running this example, make sure a view by the name 'Table View' exists.


```vb
Sub GetView() 
 
 'Returns a view called Table View 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 'Return a view called Table View 
 
 Set objView = objViews.Item("Table View") 
 
End Sub
```

The following example adds a new view of type  **olIconView** in the user's Notes folder.


> [!NOTE] 
> The  **Add** method will fail if a view with the same name already exists.




```vb
Sub CreateView() 
 
 'Creates a new view 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objNewView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderNotes).Views 
 
 Set objNewView = objViews.Add(Name:="New Icon View Type", _ 
 
 ViewType:=olIconView, SaveOption:=olViewSaveOptionThisFolderEveryone) 
 
 
 
End Sub
```

 The following example removes the above view, "New Icon View Type", from the collection.




```vb
Sub DeleteView() 
 
 'Deletes a view from the collection 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objNewView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderNotes).Views 
 
 objViews.Remove ("New Icon View Type") 
 
End Sub
```


## Events



|Name|
|:-----|
|[ViewAdd](Outlook.Views.ViewAdd.md)|
|[ViewRemove](Outlook.ViewRemove.md)|

## Methods



|Name|
|:-----|
|[Add](Outlook.Views.Add.md)|
|[Item](Outlook.Views.Item.md)|
|[Remove](Outlook.Views.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Views.Application.md)|
|[Class](Outlook.Views.Class.md)|
|[Count](Outlook.Views.Count.md)|
|[Parent](Outlook.Views.Parent.md)|
|[Session](Outlook.Views.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]