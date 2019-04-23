---
title: Application.Nz method (Access)
keywords: vbaac10.chm12554
f1_keywords:
- vbaac10.chm12554
ms.prod: access
api_name:
- Access.Application.Nz
ms.assetid: 669fe962-3881-83bb-cc40-ec9b23b44116
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.Nz method (Access)

You can use the **Nz** function to return zero, a zero-length string (" "), or another specified value when a **Variant** is **Null**. For example, you can use this function to convert a **Null** value to another value and prevent it from propagating through an expression.


## Syntax

_expression_.**Nz** (_Value_, _ValueIfNull_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**Variant**|A variable of data type **Variant**.|
| _ValueIfNull_|Optional|**Variant**|Optional (unless used in a query). A **Variant** that supplies a value to be returned if the variant argument is **Null**. This argument enables you to return a value other than zero or a zero-length string.<br/><br/>**NOTE**: If you use the **Nz** function in an expression in a query without using the _ValueIfNull_ argument, the results will be a zero-length string in the fields that contain null values. |

## Return value

Variant


## Remarks

If the _Value_ of the variant argument is **Null**, the **Nz** function returns the number zero or a zero-length string (always returns a zero-length string when used in a query expression), depending on whether the context indicates that the _Value_ should be a number or a string. If the optional _ValueIfNull_ argument is included, the **Nz** function will return the value specified by that argument if the variant argument is **Null**. When used in a query expression, the **Nz** function should always include the _ValueIfNull_ argument.

If the _Value_ of **Variant** isn't **Null**, the **Nz** function returns the _Value_ of **Variant**.

The **Nz** function is useful for expressions that may include **Null** values. To force an expression to evaluate to a non- **Null** value even when it contains a **Null** value, use the **Nz** function to return zero, a zero-length string, or a custom return value.

For example, the expression `2 + varX` will always return a **Null** value when the **Variant** `varX` is **Null**. However, `2 + Nz(varX)` returns 2.

You can often use the **Nz** function as an alternative to the **[IIf](../language/reference/user-interface-help/iif-function.md)** function. For example, in the following code, two expressions including the **IIf** function are necessary to return the desired result. The first expression including the **IIf** function is used to check the value of a variable and convert it to zero if it is **Null**.

```vb
varTemp = IIf(IsNull(varFreight), 0, varFreight) 
varResult = IIf(varTemp > 50, "High", "Low")
```

<br/>

In the next example, the **Nz** function provides the same functionality as the first expression, and the desired result is achieved in one step rather than two.

```vb
varResult = IIf(Nz(varFreight) > 50, "High", "Low")
```

<br/>

If you supply a value for the optional argument _ValueIfNull_, that value will be returned when **Variant** is **Null**. By including this optional argument, you may be able to avoid the use of an expression containing the **IIf** function. For example, the following expression uses the **IIf** function to return a string if the value of `varFreight` is **Null**.

```vb
varResult = IIf(IsNull(varFreight), "No Freight Charge", varFreight)
```

<br/>

In the next example, the optional argument supplied to the **Nz** function provides the string to be returned if `varFreight` is **Null**.

```vb
varResult = Nz(varFreight, "No Freight Charge")
```


## Example

The following example evaluates a control on a form and returns one of two strings based on the control's value. If the value of the control is **Null**, the procedure uses the **Nz** function to convert a **Null** value to a zero-length string.


```vb
Public Sub CheckValue() 
 
    Dim frm As Form 
    Dim ctl As Control 
    Dim varResult As Variant 
 
    ' Return Form object variable pointing to Orders form. 
    Set frm = Forms!Orders 
 
    ' Return Control object variable pointing to ShipRegion. 
    Set ctl = frm!ShipRegion 
 
    ' Choose result based on value of control. 
    varResult = IIf(Nz(ctl.Value) = vbNullString, _ 
        "No value.", "Value is " & ctl.Value & ".") 
 
    ' Display result. 
    MsgBox varResult, vbExclamation 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
