---
title: AutoFormatRules object (Outlook)
keywords: vbaol11.chm3210
f1_keywords:
- vbaol11.chm3210
ms.prod: outlook
api_name:
- Outlook.AutoFormatRules
ms.assetid: 74514b71-964c-f17b-4df6-e1a5c5ed2b52
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoFormatRules object (Outlook)

Represents the collection of  **[AutoFormatRule](Outlook.AutoFormatRule.md)** objects in a view.


## Remarks

Use the  **[Add](Outlook.AutoFormatRules.Add.md)** method or the **[Insert](Outlook.AutoFormatRules.Insert.md)** method of the **AutoFormatRules** collection to create a new formatting rule for the following objects derived from the **[View](Outlook.View.md)** object:


-  **[BusinessCardView](Outlook.businessCardView.md)**
    
-  **[CalendarView](Outlook.CalendarView.md)**
    
-  **[CardView](Outlook.CardView.md)**
    
-  **[IconView](Outlook.IconView.md)**
    
-  **[TableView](Outlook.TableView.md)**
    
-  **[TimelineView Object](Outlook.TimelineView.md)**
    
 **AutoFormatRule** objects contained in an **AutoFormatRules** collection are applied to each Outlook item in the order in which they are contained in the collection. Changes to **AutoFormatRule** objects are persisted only if the **[Save](Outlook.AutoFormatRules.Save.md)** method of the **AutoFormatRules** collection is called.


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


## Methods



|Name|
|:-----|
|[Add](Outlook.AutoFormatRules.Add.md)|
|[Insert](Outlook.AutoFormatRules.Insert.md)|
|[Item](Outlook.AutoFormatRules.Item.md)|
|[Remove](Outlook.AutoFormatRules.Remove.md)|
|[RemoveAll](Outlook.AutoFormatRules.RemoveAll.md)|
|[Save](Outlook.AutoFormatRules.Save.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.AutoFormatRules.Application.md)|
|[Class](Outlook.AutoFormatRules.Class.md)|
|[Count](Outlook.AutoFormatRules.Count.md)|
|[Parent](Outlook.AutoFormatRules.Parent.md)|
|[Session](Outlook.AutoFormatRules.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]