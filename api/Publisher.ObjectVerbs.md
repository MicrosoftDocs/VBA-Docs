---
title: ObjectVerbs object (Publisher)
keywords: vbapb10.chm4587519
f1_keywords:
- vbapb10.chm4587519
ms.prod: publisher
api_name:
- Publisher.ObjectVerbs
ms.assetid: e04cf7db-ee56-7d95-9f5c-7ecee1844866
ms.date: 06/01/2019
localization_priority: Normal
---


# ObjectVerbs object (Publisher)

Represents the collection of OLE verbs for the specified OLE object. OLE verbs are the operations supported by an OLE object. Commonly used OLE verbs are play and edit.
 
## Remarks

Use the **[OLEFormat.ObjectVerbs](Publisher.OLEFormat.ObjectVerbs.md)** property to return an **ObjectVerbs** object. 

## Example

The following example displays all the available verbs for the OLE object contained in the first shape on the first page in the active publication. For this example to work, the specified shape must contain an OLE object.
 
```vb
Sub GetVerbs() 
 Dim intCount As Integer 
 
 With ActiveDocument.Pages(1).Shapes(1).OLEFormat 
 For intCount = 1 To .ObjectVerbs.Count 
 MsgBox .ObjectVerbs(intCount) 
 Next 
 End With 
End Sub
```


## Methods

- [Item](Publisher.ObjectVerbs.Item.md)

## Properties

- [Application](Publisher.ObjectVerbs.Application.md)
- [Count](Publisher.ObjectVerbs.Count.md)
- [Parent](Publisher.ObjectVerbs.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]