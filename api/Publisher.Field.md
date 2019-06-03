---
title: Field object (Publisher)
keywords: vbapb10.chm6160383
f1_keywords:
- vbapb10.chm6160383
ms.prod: publisher
api_name:
- Publisher.Field
ms.assetid: 93da311a-b834-f990-60e9-786d4f6a16f1
ms.date: 05/31/2019
localization_priority: Normal
---


# Field object (Publisher)

Represents a field. The **Field** object is a member of the **[Fields](Publisher.Fields.md)** collection. The **Fields** collection represents the fields in a selection, range, or publication.
 
## Remarks

The **pbFieldPageNumber** constant is a member of the **[PbFieldType](publisher.pbfieldtype.md)** group of constants, which includes all the various field types.

Use **[TextRange.Fields](Publisher.TextRange.Fields.md)** (_index_), where _index_ is the index number, to return a single **Field** object. The index number represents the position of the field in the selection, range, or publication. 
 
## Example

The following example counts the number of fields in the active publication and displays the count in a message.

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
 MsgBox "You have " & intCount & " fields in your publication." 
 Else 
 MsgBox "You have no fields in your publication." 
 End If 
End Sub
```


## Methods

- [Unlink](Publisher.Field.Unlink.md)

## Properties

- [Application](Publisher.Field.Application.md)
- [Code](Publisher.Field.Code.md)
- [Next](Publisher.Field.Next.md)
- [Parent](Publisher.Field.Parent.md)
- [PhoneticGuide](Publisher.Field.PhoneticGuide.md)
- [Result](Publisher.Field.Result.md)
- [TextRange](Publisher.Field.TextRange.md)
- [Type](Publisher.field.type.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]