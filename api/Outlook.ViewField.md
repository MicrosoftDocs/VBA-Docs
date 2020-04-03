---
title: ViewField object (Outlook)
keywords: vbaol11.chm3205
f1_keywords:
- vbaol11.chm3205
ms.prod: outlook
api_name:
- Outlook.ViewField
ms.assetid: 997319f0-7ff3-a712-8484-2e442965e187
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewField object (Outlook)

Represents a view field, used to display information in a view.


## Remarks

Use the  **[Add](Outlook.ViewFields.Add.md)** method of the **[ViewFields](Outlook.ViewFields.md)** collection to add an Outlook item property to the following objects derived from the **[View](Outlook.View.md)** object:


-  **[CardView](Outlook.CardView.md)**
    
-  **[TableView](Outlook.TableView.md)**
    
Use the  **[ColumnFormat](Outlook.ViewField.ColumnFormat.md)** property to access the **[ColumnFormat](Outlook.ColumnFormat.md)** object representing the display properties associated with the view field. Use the **[ViewXMLSchemaName](Outlook.ViewField.ViewXMLSchemaName.md)** property to obtain the name of the view field as referenced in the XML definition of the view.


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


## Properties



|Name|
|:-----|
|[Application](Outlook.ViewField.Application.md)|
|[Class](Outlook.ViewField.Class.md)|
|[ColumnFormat](Outlook.ViewField.ColumnFormat.md)|
|[Parent](Outlook.ViewField.Parent.md)|
|[Session](Outlook.ViewField.Session.md)|
|[ViewXMLSchemaName](Outlook.ViewField.ViewXMLSchemaName.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]