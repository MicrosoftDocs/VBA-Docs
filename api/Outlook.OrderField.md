---
title: OrderField object (Outlook)
keywords: vbaol11.chm3187
f1_keywords:
- vbaol11.chm3187
ms.prod: outlook
api_name:
- Outlook.OrderField
ms.assetid: 4ae32270-bde9-3178-bca3-f8d145779d3d
ms.date: 06/08/2017
localization_priority: Normal
---


# OrderField object (Outlook)

Represents an order field, used to sort information in a view.


## Remarks

Use the  **[Add](Outlook.ViewFields.Add.md)** method of the **[OrderFields](Outlook.OrderFields.md)** object to add an Outlook item property to the **SortFields** collection for the following objects derived from the **[View](Outlook.View.md)** object:


-  **[BusinessCardView](Outlook.businessCardView.md)**
    
-  **[CardView](Outlook.CardView.md)**
    
-  **[IconView](Outlook.IconView.md)**
    
-  **[TableView](Outlook.TableView.md)**
    
Use the  **[ViewXMLSchemaName](Outlook.OrderField.ViewXMLSchemaName.md)** property to obtain the name of the order field as referenced in the XML definition of the view.

 **OrderField** objects contained in an **OrderFields** collection are applied to Outlook items displayed in the view in the order in which the objects are contained in the collection. For each **OrderField** object, use the **[IsDescending](Outlook.OrderField.IsDescending.md)** property to determine whether to sort the contents of the order field in ascending or descending order.


## Example

The following Visual Basic for Applications (VBA) example iterates through the  **[SortFields](Outlook.TableView.SortFields.md)** collection of the current **[TableView](Outlook.TableView.md)** object, displaying the label and XML schema names of each **OrderField** object in the collection.


```vb
Private Sub DisplayTableViewSortFields() 
 
 Dim objTableView As TableView 
 
 Dim objOrderField As OrderField 
 
 Dim strOutput As String 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Iterate through the OrderFields collection for 
 
 ' the table view, obtaining the label and the 
 
 ' XML schema name for each field used to sort 
 
 ' the items in the view. 
 
 For Each objOrderField In objTableView.SortFields 
 
 With objOrderField 
 
 strOutput = strOutput & .ColumnFormat.Label & _ 
 
 " (" & .ViewXMLSchemaName & ")" & vbCrLf 
 
 End With 
 
 Next 
 
 
 
 ' Display a dialog box containing the concatenated 
 
 ' sort field information. 
 
 MsgBox strOutput 
 
 End If 
 
End Sub 
 

```


## Properties



|Name|
|:-----|
|[Application](Outlook.OrderField.Application.md)|
|[Class](Outlook.OrderField.Class.md)|
|[IsDescending](Outlook.OrderField.IsDescending.md)|
|[Parent](Outlook.OrderField.Parent.md)|
|[Session](Outlook.OrderField.Session.md)|
|[ViewXMLSchemaName](Outlook.OrderField.ViewXMLSchemaName.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]