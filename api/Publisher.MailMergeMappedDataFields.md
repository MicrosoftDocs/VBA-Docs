---
title: MailMergeMappedDataFields object (Publisher)
keywords: vbapb10.chm6553599
f1_keywords:
- vbapb10.chm6553599
ms.prod: publisher
api_name:
- Publisher.MailMergeMappedDataFields
ms.assetid: 7f33bf07-9cbb-e171-d276-d5ccb06abb95
ms.date: 05/31/2019
localization_priority: Normal
---


# MailMergeMappedDataFields object (Publisher)

A collection of **[MailMergeMappedDataField](Publisher.MailMergeMappedDataField.md)** objects that represents the mapped data fields available in Microsoft Publisher.
 
## Remarks

Use the **[MappedDataFields](Publisher.MailMergeDataSource.MappedDataFields.md)** property of the **MailMergeDataSource** object to return the **MailMergeMappedDataFields** collection. 

## Example

This example creates a table on a new page of the current publication and lists the mapped data fields available in Publisher and the fields in the data source to which they are mapped. This example assumes that the current publication is a mail merge publication and that the data source fields have corresponding mapped data fields.

```vb
Sub MappedFields() 
 Dim intCount As Integer 
 Dim intRows As Integer 
 Dim docPub As Document 
 Dim pagNew As Page 
 Dim shpTable As Shape 
 Dim tblTable As Table 
 Dim rowTable As Row 
 
 On Error Resume Next 
 
 Set docPub = ThisDocument 
 Set pagNew = ThisDocument.Pages.Add(Count:=1, After:=1) 
 intRows = docPub.MailMerge.DataSource.MappedDataFields.Count + 1 
 
 'Creates new table with a heading row 
 Set shpTable = pagNew.Shapes.AddTable(NumRows:=intRows, _ 
 numColumns:=2, Left:=100, Top:=100, Width:=400, Height:=12) 
 Set tblTable = shpTable.Table 
 With tblTable.Rows(1) 
 With .Cells(1).Text 
 .Text = "Mapped Data Field" 
 .Font.Bold = msoTrue 
 End With 
 With .Cells(2).Text 
 .Text = "Data Source Field" 
 .Font.Bold = msoTrue 
 End With 
 End With 
 
 With docPub.MailMerge.DataSource 
 For intCount = 2 To intRows - 1 
 'Inserts mapped data field name and the 
 'corresponding data source field name 
 tblTable.Rows(intCount - 1).Cells(1).Text _ 
 .Text = .MappedDataFields(Index:=intCount).Name 
 tblTable.Rows(intCount - 1).Cells(2).Text _ 
 .Text = .MappedDataFields(Index:=intCount).DataFieldName 
 Next 
 End With 
End Sub
```


## Methods

- [Item](Publisher.MailMergeMappedDataFields.Item.md)

## Properties

- [Application](Publisher.MailMergeMappedDataFields.Application.md)
- [Count](Publisher.MailMergeMappedDataFields.Count.md)
- [Parent](Publisher.MailMergeMappedDataFields.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]