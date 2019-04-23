---
title: BusinessCardView.Standard property (Outlook)
keywords: vbaol11.chm2930
f1_keywords:
- vbaol11.chm2930
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.Standard
ms.assetid: b14971c1-551e-5a00-e712-46d49806dfab
ms.date: 06/08/2017
localization_priority: Normal
---


# BusinessCardView.Standard property (Outlook)

Returns a  **Boolean** value that indicates whether the **[BusinessCardView](Outlook.BusinessCardView.md)** object is a built-in Outlook view. Read-only.


## Syntax

_expression_. `Standard`

 _expression_ An expression that returns a [BusinessCardView](Outlook.BusinessCardView.md) object.


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


[BusinessCardView Object](Outlook.BusinessCardView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]