---
title: DoCmd.OpenDataAccessPage method (Access)
keywords: vbaac10.chm4648
f1_keywords:
- vbaac10.chm4648
ms.prod: access
api_name:
- Access.DoCmd.OpenDataAccessPage
ms.assetid: 130dcb88-e3e6-25a6-186c-bf541d114169
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.OpenDataAccessPage method (Access)

The **OpenDataAccessPage** method carries out the OpenDataAccessPage action in Visual Basic.


## Syntax

_expression_.**OpenDataAccessPage** (_DataAccessPageName_, _View_)

_expression_ An expression that returns a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataAccessPageName_|Required|**Variant**|A string expression that's the valid name of a data access page in the current database. If you execute Visual Basic code containing the **OpenDataAccessPage** method in a library database, Microsoft Access looks for the form with this name, first in the library database, and then in the current database.|
| _View_|Optional|**AcDataAccessPageView**|The view in which to open the data access page. In Access, this must be set to **acDataAccessPageBrowse**.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]