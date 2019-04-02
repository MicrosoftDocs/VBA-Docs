---
title: AutoFormatRule.Standard property (Outlook)
keywords: vbaol11.chm2711
f1_keywords:
- vbaol11.chm2711
ms.prod: outlook
api_name:
- Outlook.AutoFormatRule.Standard
ms.assetid: 11ba1f61-132a-11ba-529e-b38f7cb6ec57
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoFormatRule.Standard property (Outlook)

Returns a  **Boolean** value that indicates whether the **[AutoFormatRule](Outlook.AutoFormatRule.md)** object represents a built-in Outlook formatting rule. Read-only.


## Syntax

_expression_. `Standard`

_expression_ A variable that represents an [AutoFormatRule](Outlook.AutoFormatRule.md) object.


## Remarks

If the value of this property is set to  **True**, then the **[Filter](Outlook.AutoFormatRule.Filter.md)** and **[Name](Outlook.AutoFormatRule.Name.md)** properties of the **AutoFormatRule** object cannot be changed. Similarly, you cannot use the **[Remove](Outlook.AutoFormatRules.Remove.md)** method of the **[AutoFormatRules](Outlook.AutoFormatRules.md)** collection to delete a built-in Outlook formatting rule, nor can you use the **[Insert](Outlook.AutoFormatRules.Insert.md)** method of the **AutoFormatRules** collection to insert a custom formatting rule above or between the built-in Outlook formatting rules contained by that collection.


## Example

The following Visual Basic for Applications (VBA) example enumerates the  **AutoFormatRules** collection for the current **TableView** object, disabling any custom formatting rule contained by the collection.


```vb
Private Sub DisableCustomAutoFormatRules() 
 
 Dim objTableView As TableView 
 
 Dim objRule As AutoFormatRule 
 
 
 
 ' Check if the current view is a table view. 
 
 If Application.ActiveExplorer.CurrentView.ViewType = olTableView Then 
 
 
 
 ' Obtain a TableView object reference to the current view. 
 
 Set objView = Application.ActiveExplorer.CurrentView 
 
 
 
 ' Enumerate the AutoFormatRules collection for 
 
 ' the table view, disabling any custom formatting 
 
 ' rule defined for the view. 
 
 For Each objRule In objView.AutoFormatRules 
 
 If Not objRule.Standard Then 
 
 objRule.Enabled = False 
 
 End If 
 
 Next 
 
 
 
 ' Save and apply the table view. 
 
 objView.Save 
 
 objView.Apply 
 
 End If 
 
End Sub
```


## See also


[AutoFormatRule Object](Outlook.AutoFormatRule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]