---
title: Fields object (Publisher)
keywords: vbapb10.chm6094847
f1_keywords:
- vbapb10.chm6094847
ms.prod: publisher
api_name:
- Publisher.Fields
ms.assetid: fd7c95d9-bc34-95ee-180d-b99f3629eb33
ms.date: 05/31/2019
localization_priority: Normal
---


# Fields object (Publisher)

A collection of **[Field](Publisher.Field.md)** objects that represent all the fields in a text range.
 

## Remarks

The **Count** property for this collection in a publication returns the number of items in a specified shape or selection.
 
Use the **[TextRange.Fields](Publisher.TextRange.Fields.md)** property to return the **Fields** collection. Use **Fields** (_index_), where _index_ is the index number, to return a single **Field** object. The index number represents the position of the field in the selection, range, or publication. 
 

## Example

The following example displays the field code and the result of the first field in each text box in the active publication.

```vb
Sub ShowFieldCodes() 
 Dim pagPage As Page 
 Dim shpShape As Shape 
 
 For Each pagPage In ActiveDocument.Pages 
 For Each shpShape In pagPage.Shapes 
 If shpShape.Type = pbTextFrame Then 
 With shpShape.TextFrame.TextRange 
 If .Fields.Count > 0 Then 
 MsgBox "Code = " & .Fields(1).Code & vbLf _ 
 & "Result = " & .Fields(1).Result & vbLf 
 End If 
 End With 
 End If 
 Next 
 Next 
End Sub
```


## Methods

- [AddHorizontalInVertical](Publisher.Fields.AddHorizontalInVertical.md)
- [AddPhoneticGuide](Publisher.Fields.AddPhoneticGuide.md)
- [Item](Publisher.Fields.Item.md)
- [Unlink](Publisher.Fields.Unlink.md)

## Properties

- [Application](Publisher.Fields.Application.md)
- [Count](Publisher.Fields.Count.md)
- [Parent](Publisher.Fields.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]