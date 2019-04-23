---
title: Application.CreateGroupLevel method (Access)
keywords: vbaac10.chm12524
f1_keywords:
- vbaac10.chm12524
ms.prod: access
api_name:
- Access.Application.CreateGroupLevel
ms.assetid: 880c1e36-b7b5-7ea4-a2ca-d7c3f0a5a7be
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.CreateGroupLevel method (Access)

You can use the **CreateGroupLevel** method to specify a field or expression on which to group or sort data in a report.


## Syntax

_expression_.**CreateGroupLevel** (_ReportName_, _Expression_, _Header_, _Footer_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ReportName_|Required|**String**|The name of the report that will contain the new group level.|
| _Expression_|Required|**String**|The field or expression to sort or group on.|
| _Header_|Required|**Integer**|Indicates that a field or expression will have an associated group header. If the _Header_ argument is **True** (1), the field or expression will have a group header. If the _Header_ argument is **False** (0), the field or expression won't. You can create a header by setting the argument to **True**.|
| _Footer_|Required|**Integer**|Indicates a field or expression will have an associated group footer. If the _Footer_ argument is **True** (1), the field or expression will have a group footer. If the _Footer_ argument is **False** (0), the field or expression won't. You can create a footer by setting the argument to **True**.|

## Return value

Long


## Remarks

For example, suppose you are building a custom wizard that provides the user with a choice of fields on which to group data when designing a report. Call the **CreateGroupLevel** method from your wizard to create the appropriate groups according to the user's choice.

You can use the **CreateGroupLevel** method when designing a wizard that creates a report with groups or totals. The **CreateGroupLevel** method groups or sorts data on the specified field or expression and creates a header and/or footer for the group level.

The **CreateGroupLevel** method is available only in report Design view.

Microsoft Access uses the **[GroupLevel](Access.Report.GroupLevel.md)** property array to keep track of the group levels created for a report. The **CreateGroupLevel** method adds a new group level to the array, based on the _expression_ argument. The **CreateGroupLevel** method then returns an index value that represents the new group level's position in the array. The first field or expression that you sort or group on is level 0, the second is level 1, and so on. You can have up to ten group levels in a report (0 to 9).

When you specify that either the _Header_ or _Footer_ argument, or both, is **True**, the **[GroupHeader](Access.GroupLevel.GroupHeader.md)** and **[GroupFooter](Access.GroupLevel.GroupFooter.md)** properties in a report are set to Yes, and a header and/or footer is created for the group level.

After a header or footer is created, you can set other **GroupLevel** properties: **[GroupOn](Access.GroupLevel.GroupOn.md)**, **[GroupInterval](Access.GroupLevel.GroupInterval.md)**, and **[KeepTogether](Access.GroupLevel.KeepTogether.md)**.

> [!NOTE] 
> If your wizard creates group levels in a new or existing report, it must open the report in Design view.


## Example

The following example creates a group level on an **OrderDate** field on a report called OrderReport. The report on which the group level is to be created must be open in Design view. Because the _Header_ and _Footer_ arguments are set to **True** (1), the method creates both the header and footer for the group level. The header and footer are then sized.


```vb
Sub CreateGL() 
 Dim varGroupLevel As Variant 
 
 ' Create new group level on OrderDate field. 
 varGroupLevel = CreateGroupLevel("OrderReport", "OrderDate", _ 
 True, True) 
 ' Set height of header/footer sections. 
 Reports!OrderReport.Section(acGroupLevel1Header).Height = 400 
 Reports!OrderReport.Section(acGroupLevel1Footer).Height = 400 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]