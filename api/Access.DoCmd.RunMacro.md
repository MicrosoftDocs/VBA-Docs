---
title: DoCmd.RunMacro method (Access)
keywords: vbaac10.chm4175
f1_keywords:
- vbaac10.chm4175
ms.prod: access
api_name:
- Access.DoCmd.RunMacro
ms.assetid: 2abb0056-3f8a-337b-307f-6d653aa2b963
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.RunMacro method (Access)

The **RunMacro** method carries out the RunMacro action in Visual Basic.


## Syntax

_expression_.**RunMacro** (_MacroName_, _RepeatCount_, _RepeatExpression_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MacroName_|Required|**Variant**|A string expression that's the valid name of a macro in the current database. If you run Visual Basic code containing the **RunMacro** method in a library database, Microsoft Access looks for the macro with this name in the library database and doesn't look for it in the current database.|
| _RepeatCount_|Optional|**Variant**|A numeric expression that evaluates to an integer, which is the number of times the macro will run.|
| _RepeatExpression_|Optional|**Variant**| A numeric expression that's evaluated each time the macro runs. When it evaluates to **False** (0), the macro stops running.|

## Remarks

You can use the **RunMacro** method to run a macro.

You can use _MacroGroupName_._MacroName_ syntax for the _MacroName_ argument to run a particular macro in a macro group.

If you specify the _RepeatExpression_ argument and leave the _RepeatCount_ argument blank, you must include the _RepeatCount_ argument's comma. If you leave a trailing argument blank, don't use a comma following the last argument that you specify.


## Example

The following example runs the macro Print Sales that will print the sales report twice.

```vb
DoCmd.RunMacro "Print Sales", 2
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
