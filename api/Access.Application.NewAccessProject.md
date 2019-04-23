---
title: Application.NewAccessProject method (Access)
keywords: vbaac10.chm12580
f1_keywords:
- vbaac10.chm12580
ms.prod: access
api_name:
- Access.Application.NewAccessProject
ms.assetid: e3b3b9ef-31f8-885c-5c92-d269b824fbdb
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.NewAccessProject method (Access)

You can use the **NewAccessProject** method to create and open a new Microsoft Access project (.adp) as the current Access project in the Access window.


## Syntax

_expression_.**NewAccessProject** (_filepath_, _Connect_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _filepath_|Required|**String**|The name of the new Access project, including the path name and the file name extension.|
| _Connect_|Optional|**Variant**|The connection string for the Access project. See the ADO **[ConnectionString](https://docs.microsoft.com/office/client-developer/access/desktop-database-reference/connectionstring-property-ado)** property for details about this string.|

## Return value

Nothing


## Remarks

The **NewAccessProject** method enables you to create a new Access project from within Access or another application through Automation, formally called OLE Automation. For example, you can use the **NewAccessProject** method from Microsoft Excel to create a new Access project in the Access window. After you have created an instance of Access from another application, you must also create a new Access project. This Access project opens in the Access window.

If the Access project identified by _projname_ already exists, an error occurs.

The new Access project is opened under the Admin user account.

> [!NOTE] 
> To open an Access database, use the **[NewCurrentDatabase](Access.Application.NewCurrentDatabase.md)** method of the **Application** object.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]