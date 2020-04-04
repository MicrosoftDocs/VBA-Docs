---
title: TableView.AutoPreviewFont property (Outlook)
keywords: vbaol11.chm2535
f1_keywords:
- vbaol11.chm2535
ms.prod: outlook
api_name:
- Outlook.TableView.AutoPreviewFont
ms.assetid: 988e7bc4-9957-f611-b89e-1eb7a14fbfcc
ms.date: 06/08/2017
localization_priority: Normal
---


# TableView.AutoPreviewFont property (Outlook)

Returns a **[ViewFont](Outlook.ViewFont.md)** object that represents the font used when automatically previewing Outlook items in the **[TableView](Outlook.TableView.md)** object. Read-only.


## Syntax

_expression_. `AutoPreviewFont`

_expression_ A variable that represents a [TableView](Outlook.TableView.md) object.


## Example

The following Visual Basic for Applications (VBA) sample decrements the value of the  **[Size](Outlook.ViewFont.Size.md)** property for the **ViewFont** object returned from the **AutoPreviewFont** property for the current **TableView** object.


```vb
Private Sub ReduceAutoPreviewFontSize() 
 
 Dim objTableView As TableView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Decrement the Size property of the 
 
 ' ViewFont object obtained from the 
 
 ' AutoPreviewFont property, but only 
 
 ' if the font is 6 points or larger. 
 
 If objTableView.AutoPreviewFont.Size > 5 Then 
 
 objTableView.AutoPreviewFont.Size = _ 
 
 objTableView.AutoPreviewFont.Size - 1 
 
 
 
 ' Save the table view. 
 
 objTableView.Save 
 
 End If 
 
 End If 
 
End Sub
```


## See also


[TableView Object](Outlook.TableView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]