---
title: Report.GroupLevel property (Access)
keywords: vbaac10.chm13790
f1_keywords:
- vbaac10.chm13790
ms.prod: access
api_name:
- Access.Report.GroupLevel
ms.assetid: 8a40502d-84ac-0652-8c07-c4c155ec1242
ms.date: 03/15/2019
localization_priority: Normal
---


# Report.GroupLevel property (Access)

You can use the **GroupLevel** property in Visual Basic to refer to the group level that you are grouping or sorting on in a report. Read-only **GroupLevel** object.


## Syntax

_expression_.**GroupLevel** (_Index_)

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The group level, starting with 0. The first field or expression that you group on is group level 0, the second is group level 1, and so on.|

## Remarks

The following sample settings show how you use the **GroupLevel** property to refer to a group level.

|Group level|Refers to|
|:-----|:-----|
|**GroupLevel** (0)|The first field or expression that you sort or group on.|
|**GroupLevel** (1)|The second field or expression that you sort or group on.|
|**GroupLevel** (2)|The third field or expression that you sort or group on.|

The **GroupLevel** property setting is an array in which each entry identifies a group level. You can have up to 10 group levels (0 to 9).

> [!NOTE] 
> You can use this property only by using Visual Basic to set the **[SortOrder](Access.GroupLevel.SortOrder.md)**, **[GroupOn](Access.GroupLevel.GroupOn.md)**, **[GroupInterval](Access.GroupLevel.GroupInterval.md)**, **[KeepTogether](Access.GroupLevel.KeepTogether.md)**, and **ControlSource** properties. You set these properties in the **[Open](Access.Report.Open.md)** event procedure of a report.

In reports, you can group or sort on more than one field or expression. Each field or expression that you group or sort on is a group level.

You specify the fields and expressions to sort and group on by using the **[CreateGroupLevel](Access.Application.CreateGroupLevel.md)** method.

If a group is already defined for a report (the **GroupLevel** property is set to 0), you can use the **ControlSource** property to change the group level in the report's **Open** event procedure.


## Example

The following code changes the **ControlSource** property to a value contained in the **txtPromptYou** text box on the open form named **SortForm**.

```vb
Private Sub Report_Open(Cancel As Integer) 
 Me.GroupLevel(0).ControlSource _ 
 = Forms!SortForm!txtPromptYou 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]