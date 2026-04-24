---
title: OutputAllFields property
ROBOTS: INDEX
keywords: vbaac10.chm4453
f1_keywords:
- vbaac10.chm4453
api_name:
- Access.OutputAllFields
ms.assetid: b4f6e831-f82c-0850-d736-1bbda29d9f89
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# OutputAllFields property

**Applies to:** Access 2013 | Access 2016

Use the **OutputAllFields** property to show all fields in the query's underlying data source and in the field list of a form or report. Setting this property is an easy way to show all fields without having to click the Show box in the query design grid for each field in the query.

> [!NOTE] 
> The **OutputAllFields** property applies only to append, make-table, and select queries.


## Setting

The **OutputAllFields** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|Yes|Displays all the fields in the underlying tables and in the field list of a form or report.|
|No|(Default) Displays only fields that have the Show box selected in the query design grid.|

You can set this property only by using the query's property sheet.

> [!NOTE] 
> The use of an asterisk (*) in an SQL statement in place of a field name is the equivalent of setting the **OutputAllFields** property to Yes.


## Remarks

When the **OutputAllFields** property is set to Yes, the only fields you need to include in the query design grid are those that you want to sort on or specify criteria for.

When you save a filter as a query, Microsoft Access sets the **OutputAllFields** property to Yes.

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](https://learn.microsoft.com/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]