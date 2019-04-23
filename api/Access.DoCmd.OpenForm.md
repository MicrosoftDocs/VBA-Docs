---
title: DoCmd.OpenForm method (Access)
keywords: vbaac10.chm4160
f1_keywords:
- vbaac10.chm4160
ms.prod: access
api_name:
- Access.DoCmd.OpenForm
ms.assetid: a1c9d3a9-2af8-c30a-acb0-6428c70dcdb0
ms.date: 03/07/2019
localization_priority: Priority
---


# DoCmd.OpenForm method (Access)

The **OpenForm** method carries out the OpenForm action in Visual Basic.


## Syntax

_expression_.**OpenForm** (_FormName_, _View_, _FilterName_, _WhereCondition_, _DataMode_, _WindowMode_, _OpenArgs_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FormName_|Required|**Variant**|A string expression that's the valid name of a form in the current database. If you execute Visual Basic code containing the **OpenForm** method in a library database, Access looks for the form with this name first in the library database, and then in the current database.|
| _View_|Optional|**[AcFormView](Access.AcFormView.md)**|An **AcFormView** constant that specifies the view in which the form will open. The default value is **acNormal**.|
| _FilterName_|Optional|**Variant**|A string expression that's the valid name of a query in the current database.|
| _WhereCondition_|Optional|**Variant**|A string expression that's a valid SQL WHERE clause without the word WHERE.|
| _DataMode_|Optional|**[AcFormOpenDataMode](Access.AcFormOpenDataMode.md)**|An **AcFormOpenDataMode** constant that specifies the data entry mode for the form. This applies only to forms opened in Form view or Datasheet view. The default value is **acFormPropertySettings**.|
| _WindowMode_|Optional|**[AcWindowMode](Access.AcWindowMode.md)**|An **AcWindowMode** constant that specifies the window mode in which the form opens. The default value is **acWindowNormal**.|
| _OpenArgs_|Optional|**Variant**|A string expression. This expression is used to set the form's **OpenArgs** property. This setting can then be used by code in a form module, such as the **Open** event procedure. The **OpenArgs** property can also be referred to in macros and expressions.<br/><br/>For example, suppose that the form that you open is a continuous-form list of clients. If you want the focus to move to a specific client record when the form opens, you can specify the client name with the _OpenArgs_ argument, and then use the **FindRecord** method to move the focus to the record for the client with the specified name.|

## Remarks

You can use the **OpenForm** method to open a form in Form view, form Design view, Print Preview, or Datasheet view. You can select data entry and window modes for the form and restrict the records that the form displays.

The maximum length of the _WhereCondition_ argument is 32,768 characters (unlike the _WhereCondition_ action argument in the Macro window, whose maximum length is 256 characters).


## Example

The following example opens the **Employees** form in Form view and displays only records with King in the **LastName** field. The displayed records can be edited, and new records can be added.

```vb
DoCmd.OpenForm "Employees", , ,"LastName = 'King'"
```

<br/>

The following example opens the **frmMainEmployees** form in Form view and displays only records that apply to the department chosen in the **cboDept** combo box. The displayed records can be edited, and new records can be added.


```vb
Private Sub cmdFilter_Click()
    DoCmd.OpenForm "frmMainEmployees", , , "DepartmentID=" & cboDept.Value
End Sub
```

<br/>

The following example shows how to use the _WhereCondition_ argument of the **OpenForm** method to filter the records displayed on a form as it is opened.

```vb
Private Sub cmdShowOrders_Click()
If Not Me.NewRecord Then
    DoCmd.OpenForm "frmOrder", _
        WhereCondition:="CustomerID=" & Me.txtCustomerID
End If
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
