---
title: IconView.Standard property (Outlook)
keywords: vbaol11.chm2570
f1_keywords:
- vbaol11.chm2570
ms.prod: outlook
api_name:
- Outlook.IconView.Standard
ms.assetid: 13816c3b-a35f-30cf-c63e-fb7d52a0a942
ms.date: 06/08/2017
localization_priority: Normal
---


# IconView.Standard property (Outlook)

Returns a  **Boolean** value that indicates whether the **[IconView](Outlook.IconView.md)** object is a built-in Outlook view. Read-only.


## Syntax

_expression_. `Standard`

_expression_ A variable that represents an [IconView](Outlook.IconView.md) object.


## Remarks

The  **[Reset](Outlook.View.Reset.md)** method can only be used on a view if the value of this property is set to **True**.


## Example

The following Visual Basic for Applications (VBA) example enumerates through the  **[Views](Outlook.Views.md)** collection of the current **[Folder](Outlook.Folder.md)** object, using the **Standard** property to determine if a **View** object is a built-in Outlook view. If the **View** object is a built-in Outlook view, the sample calls the **Reset** method to reset the view to its default settings. Otherwise, the sample uses the **[Delete](Outlook.View.Delete.md)** method to delete the view.


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


[IconView Object](Outlook.IconView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]