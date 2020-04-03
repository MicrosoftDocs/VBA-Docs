---
title: ViewFields object (Outlook)
keywords: vbaol11.chm3184
f1_keywords:
- vbaol11.chm3184
ms.prod: outlook
api_name:
- Outlook.ViewFields
ms.assetid: 2516faed-ed11-6cb3-ce9c-b6afa788e909
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewFields object (Outlook)

Represents the collection of  **[ViewField](Outlook.ViewField.md)** objects in a view.


## Remarks

The  **ViewFields** collection represents the Outlook item properties available for display in the view. Use the **[Add](Outlook.ViewFields.Add.md)** method of the **ViewFields** collection to add a view field for the following objects derived from the **[View](Outlook.View.md)** object:


-  **[CardView](Outlook.CardView.md)**
    
-  **[TableView](Outlook.TableView.md)**
    
In a table view, the order of  **ViewField** objects in the **ViewFields** collection is not the same as the order that field columns are displayed in the table view. A workaround to obtain the column order is to parse the string returned by the **[View.XML](Outlook.View.XML.md)** property.


## Example

The following Visual Basic for Applications (VBA) example iterates through the  **[ViewFields](Outlook.TableView.ViewFields.md)** collection of the current **[TableView](Outlook.TableView.md)** object, displaying the label and XML schema names of each **ViewField** object in the collection.


```vb
Private Sub DisplayTableViewFields() 
 
 Dim objTableView As TableView 
 
 Dim objViewField As ViewField 
 
 Dim strOutput As String 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Iterate through the ViewFields collection for 
 
 ' the table view, obtaining the label and the 
 
 ' XML schema name for each field included in 
 
 ' the view. 
 
 For Each objViewField In objTableView.ViewFields 
 
 With objViewField 
 
 strOutput = strOutput & .ColumnFormat.Label & _ 
 
 " (" & .ViewXMLSchemaName & ")" & vbCrLf 
 
 End With 
 
 Next 
 
 
 
 ' Display a dialog box containing the concatenated 
 
 ' view field information. 
 
 MsgBox strOutput 
 
 End If 
 
End Sub 
 

```


## Methods



|Name|
|:-----|
|[Add](Outlook.ViewFields.Add.md)|
|[Insert](Outlook.ViewFields.Insert.md)|
|[Item](Outlook.ViewFields.Item.md)|
|[Remove](Outlook.ViewFields.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.ViewFields.Application.md)|
|[Class](Outlook.ViewFields.Class.md)|
|[Count](Outlook.ViewFields.Count.md)|
|[Parent](Outlook.ViewFields.Parent.md)|
|[Session](Outlook.ViewFields.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]