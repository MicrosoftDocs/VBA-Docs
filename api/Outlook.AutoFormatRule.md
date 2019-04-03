---
title: AutoFormatRule object (Outlook)
keywords: vbaol11.chm3209
f1_keywords:
- vbaol11.chm3209
ms.prod: outlook
api_name:
- Outlook.AutoFormatRule
ms.assetid: 6d295c41-17f9-8e67-4595-4330fd3cec99
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoFormatRule object (Outlook)

Represents a formatting rule used by a  **[View](Outlook.View.md)** object to determine how to format Outlook items displayed within that view.


## Remarks

Use the  **[Add](Outlook.AutoFormatRules.Add.md)** method or the **[Insert](Outlook.AutoFormatRules.Insert.md)** method of the **[AutoFormatRules](Outlook.AutoFormatRules.md)** collection to create a new formatting rule for the following objects:


-  **[CalendarView](Outlook.CalendarView.md)**
    
-  **[CardView](Outlook.CardView.md)**
    
-  **[TableView](Outlook.TableView.md)**
    

### Built-In and Custom Formatting Rules

Microsoft Outlook provides a set of built-in formatting rules that can be disabled but cannot be removed or reordered. Custom formatting rules, defined either programmatically or by user action, cannot be moved above or between built-in formatting rules. Use the  **[Standard](Outlook.AutoFormatRule.Standard.md)** property to determine whether a formatting rule is built-in or custom.


### Applying Formatting Rules

Formatting rules are checked and applied against each Outlook item, in the order in which they are contained within the  **AutoFormatRules** collection. Use the **[Enabled](Outlook.AutoFormatRule.Enabled.md)** property to enable or disable a formatting rule, the **[Filter](Outlook.AutoFormatRule.Filter.md)** property to define the conditions an Outlook item must meet to be formatted by the formatting rule, and the **[Font](Outlook.AutoFormatRule.Font.md)** property to specify the format to be applied by the formatting rule.


## Example

The following Visual Basic for Applications (VBA) example enumerates the  **[AutoFormatRules](Outlook.TableView.AutoFormatRules.md)** collection for the current **TableView** object, disabling any custom formatting rule contained by the collection.


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


## Properties



|Name|
|:-----|
|[Application](Outlook.AutoFormatRule.Application.md)|
|[Class](Outlook.AutoFormatRule.Class.md)|
|[Enabled](Outlook.AutoFormatRule.Enabled.md)|
|[Filter](Outlook.AutoFormatRule.Filter.md)|
|[Font](Outlook.AutoFormatRule.Font.md)|
|[Name](Outlook.AutoFormatRule.Name.md)|
|[Parent](Outlook.AutoFormatRule.Parent.md)|
|[Session](Outlook.AutoFormatRule.Session.md)|
|[Standard](Outlook.AutoFormatRule.Standard.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]