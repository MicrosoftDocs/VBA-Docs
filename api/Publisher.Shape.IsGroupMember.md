---
title: Shape.IsGroupMember property (Publisher)
keywords: vbapb10.chm2228337
f1_keywords:
- vbapb10.chm2228337
ms.prod: publisher
api_name:
- Publisher.Shape.IsGroupMember
ms.assetid: bbd9b662-b47d-d5cf-6858-e208c44f88a0
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.IsGroupMember property (Publisher)

Returns **True** if the specified shape is a member of a group; otherwise, **False**. Read-only **Boolean**.


## Syntax

_expression_.**IsGroupMember**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Return value

Boolean


## Remarks

The object returned by the **ParentGroupShape** property can be used to determine the parent shape for the group.


## Example

The following statement can be used to return a **True** value if the first shape of the active publication is a group member.

```vb
blnGrouped = Application.ActiveDocument.MasterPages _ 
 .Item.Shapes(1).IsGroupMember
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]