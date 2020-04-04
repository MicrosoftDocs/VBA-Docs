---
title: TableView.GridLineStyle property (Outlook)
keywords: vbaol11.chm2528
f1_keywords:
- vbaol11.chm2528
ms.prod: outlook
api_name:
- Outlook.TableView.GridLineStyle
ms.assetid: b3a95e34-4d22-e208-255d-43fc2456f0e6
ms.date: 06/08/2017
localization_priority: Normal
---


# TableView.GridLineStyle property (Outlook)

Returns or sets an **[OlGridLineStyle](Outlook.OlGridLineStyle.md)** constant that represents the line style used for grid lines in the **[TableView](Outlook.TableView.md)** object. Read/write.


## Syntax

_expression_. `GridLineStyle`

_expression_ A variable that represents a [TableView](Outlook.TableView.md) object.


## Example

The following Visual Basic for Applications (VBA) example sets the  **GridLineStyle** property of the current **TableView** object to display the grid with small dotted lines.


```vb
Private Sub SetDottedGridLines() 
 
 Dim objTableView As TableView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 With objTableView 
 
 ' Set the GridLineStyle property so that 
 
 ' the grid in the table view are 
 
 ' displayed with thin dotted lines. 
 
 .GridLineStyle = olGridLineSmallDots 
 
 
 
 ' Save the table view. 
 
 .Save 
 
 End With 
 
 End If 
 
End Sub
```


## See also


[TableView Object](Outlook.TableView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]