---
title: GroupLevel object (Access)
keywords: vbaac10.chm12247
f1_keywords:
- vbaac10.chm12247
ms.prod: access
api_name:
- Access.GroupLevel
ms.assetid: fdc4f24e-98aa-27bd-7a9d-271d48912dfa
ms.date: 03/20/2019
localization_priority: Normal
---


# GroupLevel object (Access)

You can use the **GroupLevel** property in Visual Basic to refer to the group level that you are grouping or sorting on in a report.


## Remarks

The **GroupLevel** property setting is an array in which each entry identifies a group level. To refer to a group level, use this syntax:

**GroupLevel** (_n_)

The number _n_ is the group level, starting with 0. The first field or expression that you group on is group level 0, the second is group level 1, and so on. You can have up to 10 group levels (0 to 9).

The following sample settings show how you use the **GroupLevel** property to refer to a group level.

|Group level|Refers to|
|:-----|:-----|
|**GroupLevel** (0)|The first field or expression that you sort or group on.|
|**GroupLevel** (1)|The second field or expression that you sort or group on.|
|**GroupLevel** (2)|The third field or expression that you sort or group on.|

You can use this property only by using Visual Basic to set the **SortOrder**, **GroupOn**, **GroupInterval**, **KeepTogether**, and **ControlSource** properties. You set these properties in the **[Open](access.report.open.md)** event procedure of a report.

In reports, you can group or sort on more than one field or expression. Each field or expression that you group or sort on is a group level.

You specify the fields and expressions to sort and group on by using the **[CreateGroupLevel](access.application.creategrouplevel.md)** method.

If a group is already defined for a report (the **GroupLevel** property is set to 0), you can use the **ControlSource** property to change the group level in the report's **Open** event procedure. 

For example, the following code changes the **ControlSource** property to a value contained in the **txtPromptYou** text box on the open form named **SortForm**.

```vb
Private Sub Report_Open(Cancel As Integer) 
 Me.GroupLevel(0).ControlSource _ 
 = Forms!SortForm!txtPromptYou 
End Sub
```


## Properties

- [Application](Access.GroupLevel.Application.md)
- [ControlSource](Access.GroupLevel.ControlSource.md)
- [GroupFooter](Access.GroupLevel.GroupFooter.md)
- [GroupHeader](Access.GroupLevel.GroupHeader.md)
- [GroupInterval](Access.GroupLevel.GroupInterval.md)
- [GroupOn](Access.GroupLevel.GroupOn.md)
- [KeepTogether](Access.GroupLevel.KeepTogether.md)
- [Parent](Access.GroupLevel.Parent.md)
- [Properties](Access.GroupLevel.Properties.md)
- [SortOrder](Access.GroupLevel.SortOrder.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]