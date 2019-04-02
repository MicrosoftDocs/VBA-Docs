---
title: TableView.RowFont property (Outlook)
keywords: vbaol11.chm2533
f1_keywords:
- vbaol11.chm2533
ms.prod: outlook
api_name:
- Outlook.TableView.RowFont
ms.assetid: 691be8dc-8811-64d0-7473-93a0fe8b4749
ms.date: 06/08/2017
localization_priority: Normal
---


# TableView.RowFont property (Outlook)

Returns a  **[ViewFont](Outlook.ViewFont.md)** object that represents the font used when displaying rows in the **[TableView](Outlook.TableView.md)** object. Read-only.


## Syntax

_expression_. `RowFont`

_expression_ A variable that represents a [TableView](Outlook.TableView.md) object.


## Example

The following Visual Basic for Applications (VBA) sample increments the value of the  **[Size](Outlook.ViewFont.Size.md)** property for the **ViewFont** object returned from the **RowFont** property for the current **TableView** object.


```vb
Private Sub IncreaseRowFontSize() 
 
 Dim objTableView As TableView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Increment the Size property of the 
 
 ' ViewFont object obtained from the 
 
 ' RowFont property, but only 
 
 ' if the font is less than 24 points 
 
 ' in size. 
 
 If objTableView.RowFont.Size < 24 Then 
 
 objTableView.RowFont.Size = _ 
 
 objTableView.RowFont.Size + 1 
 
 
 
 ' Save the table view. 
 
 objTableView.Save 
 
 End If 
 
 End If 
 
End Sub
```


## See also


[TableView Object](Outlook.TableView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]