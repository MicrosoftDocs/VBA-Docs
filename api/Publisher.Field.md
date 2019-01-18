---
title: Field Object (Publisher)
keywords: vbapb10.chm6160383
f1_keywords:
- vbapb10.chm6160383
ms.prod: publisher
api_name:
- Publisher.Field
ms.assetid: 93da311a-b834-f990-60e9-786d4f6a16f1
ms.date: 06/08/2017
localization_priority: Normal
---


# Field Object (Publisher)

Represents a field. The  **Field** object is a member of the **[Fields](Publisher.Fields.md)** collection. The **Fields** collection represents the fields in a selection, range, or publication.
 


## Remarks

The  **pbFieldPageNumber** constant is a member of the **PbFieldType** group of constants, which includes all the various field types.
 

 

## Example

Use  **[Fields](Publisher.TextRange.Fields.md)** (index), where index is the index number, to return a single **Field** object. The index number represents the position of the field in the selection, range, or publication. The following counts the number of fields in the active publication and displays the count in a message.
 

 

```vb
Sub CountFields() 
 Dim pagPage As Page 
 Dim shpShape As Shape 
 Dim fldField As Field 
 Dim intFields As Integer 
 Dim intCount As Integer 
 
 For Each pagPage In ActiveDocument.Pages 
 For Each shpShape In pagPage.Shapes 
 If shpShape.Type = pbTextFrame Then 
 intCount = intCount + shpShape.TextFrame.TextRange.Fields.Count 
 End If 
 Next 
 Next 
 If intCount > 0 Then 
 MsgBox "You have " &amp; intCount &amp; " fields in your publication." 
 Else 
 MsgBox "You have no fields in your publication." 
 End If 
End Sub
```


## Methods



|Name|
|:-----|
|[Unlink](Publisher.Field.Unlink.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.Field.Application.md)|
|[Code](Publisher.Field.Code.md)|
|[Next](Publisher.Field.Next.md)|
|[Parent](Publisher.Field.Parent.md)|
|[PhoneticGuide](Publisher.Field.PhoneticGuide.md)|
|[Result](Publisher.Field.Result.md)|
|[TextRange](Publisher.Field.TextRange.md)|
|[Type](Publisher.field.type.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]