---
title: ColumnFormat object (Outlook)
keywords: vbaol11.chm3189
f1_keywords:
- vbaol11.chm3189
ms.prod: outlook
api_name:
- Outlook.ColumnFormat
ms.assetid: acbbdd97-e695-d1e7-c7ba-24f75efbf22c
ms.date: 06/08/2017
localization_priority: Normal
---


# ColumnFormat object (Outlook)

Represents the display properties of an order field or view field in a view.


## Remarks

The  **ColumnFormat** object represents the display properties, such as the alignment or field type, of an **[OrderField](Outlook.OrderField.md)** or **[ViewField](Outlook.ViewField.md)** object. Use the **[ColumnFormat](Outlook.ViewField.ColumnFormat.md)** property of the **ViewField** object to access the display properties of a view field.

Use the  **[Label](Outlook.ColumnFormat.Label.md)** property to obtain or change the text used to label the field, or the **[Align](Outlook.ColumnFormat.Align.md)** property to determine the alignment of the contents within the field.

Use the  **[FieldType](Outlook.ColumnFormat.FieldType.md)** property to determine the type and form of the data displayed for that field, and the **[FieldFormat](Outlook.ColumnFormat.FieldFormat.md)** property to determine how to format the data for that field.


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
|[Align](Outlook.ColumnFormat.Align.md)|
|[Application](Outlook.ColumnFormat.Application.md)|
|[Class](Outlook.ColumnFormat.Class.md)|
|[FieldFormat](Outlook.ColumnFormat.FieldFormat.md)|
|[FieldType](Outlook.ColumnFormat.FieldType.md)|
|[Label](Outlook.ColumnFormat.Label.md)|
|[Parent](Outlook.ColumnFormat.Parent.md)|
|[Session](Outlook.ColumnFormat.Session.md)|
|[Width](Outlook.ColumnFormat.Width.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]