---
title: Section.KeepTogether property (Access)
keywords: vbaac10.chm12194
f1_keywords:
- vbaac10.chm12194
ms.prod: access
api_name:
- Access.Section.KeepTogether
ms.assetid: dbe3780b-2150-4b4c-d8bf-5685ab48181e
ms.date: 03/23/2019
localization_priority: Normal
---


# Section.KeepTogether property (Access)

You can use the **KeepTogether** property for a section to print a form or report section all on one page. For example, you might have a group of related information that you don't want printed across two pages. The **KeepTogether** property applies only to form and report sections (except page headers and page footers). Read/write **Boolean**.


## Syntax

_expression_.**KeepTogether**

_expression_ A variable that represents a **[Section](Access.Section.md)** object.


## Remarks

The **KeepTogether** property for a section uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|Microsoft Access starts printing the section at the top of the next page if it can't print the entire section on the current page.|
|No|**False**|(Default) Access prints as much of the section as possible on the current page and prints the rest on the next page.|

You can set the **KeepTogether** property for a section only in form Design view or report Design view.

Usually, when a page break occurs while a section is being printed, Access continues printing the section on the next page. By using the section's **KeepTogether** property, you can print the section all on one page. If a section is longer than one page, Access starts printing it on the next page and continues on the following page.

If the **KeepTogether** property for a group is set to Whole Group or With First Detail, and the **KeepTogether** property for a section is set to No, the **KeepTogether** property setting for the section is ignored.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]