---
title: Form.InputParameters property (Access)
keywords: vbaac10.chm13487
f1_keywords:
- vbaac10.chm13487
ms.prod: access
api_name:
- Access.Form.InputParameters
ms.assetid: fc3e17a7-f62a-a6bb-c44a-f3a9d7efe6ac
ms.date: 03/13/2019
localization_priority: Normal
---


# Form.InputParameters property (Access)

You can use the **InputParameters** property to specify or determine the input parameters that are passed to a SQL statement in the **RecordSource** property of a form or report, or a stored procedure when used as the record source within a Microsoft Access project (.adp). Read/write **String**.


## Syntax

_expression_.**InputParameters**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

### When used with a RecordSource property

An example **InputParameters** property string used with a SQL statement in the **RecordSource** property would use the following syntax.

```vb
state char=[Forms]![formname]![StateList], salesyear smallint=[Forms]![formname]![Enter year of interest]

```

This would result in the _state_ parameter being set to the current value of the **StateList** control, and the user getting prompted for the _salesyear_ parameter. If there were any other parameters that were not in this list, they would get default values assigned.

The query should be executed with one marker for each non-default parameter in the **InputParameters** list.

A refresh or requery command (via menu, keyboard, or navigation bar) in Access should trigger a reexecute of the query. Users can do this in code by calling the standard **Requery** method. If the value of a parameter is bound to a control on the form, the current value of the control is used at requery time. The query is not automatically reexecuted when the value of the control changes.

### When used with a stored procedure

An example **InputParameters** property string used with a stored procedure would use the following syntax.

```vb
@state char=[Forms]![formname]![StateList], @salesyear smallint=[Forms]![formname]![Enter year of interest]

```

This would result in the @state parameter being set to the current value of the **StateList** control, and the user getting prompted for the @salesyear parameter. If there were any other parameters to the stored procedure that were not in this list, they would get default values assigned.

The stored procedure should be executed by using a command string containing the {call } syntax with one marker for each non-default parameter in the **InputParameters** list.

A refresh or requery command (via menu, keyboard, or navigation bar) in Access should trigger a reexecute of the stored procedure. Users can do this in code by calling the standard **Requery** method. If the value of a parameter is bound to a control on the form, the current value of the control is used at requery time. The stored procedure is not automatically reexecuted when the value of the control changes.

This builder dialog is invoked when a stored procedure is first selected as the record source of a form if the stored procedure has any parameters. After initial creation of the **InputParameters** string, this same dialog is used as a builder for changing the string. In this case, however, the list of parameters comes from what already exists in the string.

Parameter values are also settable in code by using the ActiveX Data Object's (ADO) **Command** and **Parameter** objects. If the result returns a result set, a form can be bound to it by setting the form's **Recordset** property. ADO coding is the only way to handle stored procedures that do not return result sets such as action queries, those that return output parameters, or those that return multiple result sets.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]