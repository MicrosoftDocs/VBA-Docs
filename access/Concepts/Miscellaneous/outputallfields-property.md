---
title: OutputAllFields property
ROBOTS: INDEX
keywords: vbaac10.chm4453
f1_keywords:
- vbaac10.chm4453
ms.prod: access
api_name:
- Access.OutputAllFields
ms.assetid: b4f6e831-f82c-0850-d736-1bbda29d9f89
ms.date: 06/08/2017
localization_priority: Normal
---


# OutputAllFields property

**Applies to:** Access 2013 | Access 2016

You can use the **OutputAllFields** property to show all fields in the query's underlying data source and in the field list of a form or report. Setting this property is an easy way to show all fields without having to click the Show box in the query design grid for each field in the query.

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

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]