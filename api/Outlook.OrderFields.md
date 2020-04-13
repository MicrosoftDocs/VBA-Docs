---
title: OrderFields object (Outlook)
keywords: vbaol11.chm3186
f1_keywords:
- vbaol11.chm3186
ms.prod: outlook
api_name:
- Outlook.OrderFields
ms.assetid: e115fb80-352d-fd2e-c1c3-d266776fe122
ms.date: 06/08/2017
localization_priority: Normal
---


# OrderFields object (Outlook)

Represents the collection of  **[OrderField](Outlook.OrderField.md)** objects in a view.


## Remarks

The **OrderFields** collection represents the Outlook item properties used to sort Outlook items displayed in the view. Use the **[Add](Outlook.OrderFields.Add.md)** method or the **OrderFields** collection to create a new order field for the following objects derived from the **[View](Outlook.View.md)** object:


-  **[BusinessCardView](Outlook.businessCardView.md)**
    
-  **[CardView](Outlook.CardView.md)**
    
-  **[IconView](Outlook.IconView.md)**
    
-  **[PeopleView](Outlook.peopleview.md)**
    
-  **[TableView](Outlook.TableView.md)**
    
 **OrderField** objects contained in an **OrderFields** collection are applied to Outlook items displayed in the view in the order in which the objects are contained in the collection.


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


## Methods



|Name|
|:-----|
|[Add](Outlook.OrderFields.Add.md)|
|[Insert](Outlook.OrderFields.Insert.md)|
|[Item](Outlook.OrderFields.Item.md)|
|[Remove](Outlook.OrderFields.Remove.md)|
|[RemoveAll](Outlook.OrderFields.RemoveAll.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.OrderFields.Application.md)|
|[Class](Outlook.OrderFields.Class.md)|
|[Count](Outlook.OrderFields.Count.md)|
|[Parent](Outlook.OrderFields.Parent.md)|
|[Session](Outlook.OrderFields.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]