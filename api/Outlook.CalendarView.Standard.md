---
title: CalendarView.Standard property (Outlook)
keywords: vbaol11.chm2620
f1_keywords:
- vbaol11.chm2620
ms.prod: outlook
api_name:
- Outlook.CalendarView.Standard
ms.assetid: 7d4ac52a-8a3d-25b1-6900-3799fe0fde70
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarView.Standard property (Outlook)

Returns a  **Boolean** value that indicates whether the **[CalendarView](Outlook.CalendarView.md)** object is a built-in Outlook view. Read-only.


## Syntax

_expression_. `Standard`

 _expression_ An expression that returns a [CalendarView](Outlook.CalendarView.md) object.


## Remarks

The  **[Reset](Outlook.View.Reset.md)** method can only be used on a view if the value of this property is set to **True**.


## Example

The following Visual Basic for Applications (VBA) example enumerates through the  **[Views](Outlook.Views.md)** collection of the current **[Folder](Outlook.Folder.md)** object using the **Standard** property to determine if a **View** object is a built-in Outlook view. If the **View** object is a built-in Outlook view, the sample calls the **Reset** method to reset the view to its default settings. Otherwise, the sample uses the **[Delete](Outlook.View.Delete.md)** method to delete the view.


```vb
Private Sub RemoveAllViewCustomization() 
 
 Dim objView As View 
 
 
 
 ' Enumerate each View object in the Views collection 
 
 ' of the current Folder object. 
 
 For Each objView In Application.ActiveExplorer.CurrentFolder.Views 
 
 ' If the View object is a built-in Outlook view, reset 
 
 ' the view to its default settings. If the View object 
 
 ' is a custom view, delete it. 
 
 If objView.Standard Then 
 
 objView.Reset 
 
 Else 
 
 objView.Delete 
 
 End If 
 
 Next 
 
End Sub
```


## See also


[CalendarView Object](Outlook.CalendarView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]