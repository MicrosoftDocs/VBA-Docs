---
title: Edit Watch dialog box
keywords: vbui6.chm181027
f1_keywords:
- vbui6.chm181027
ms.prod: office
ms.assetid: 2847eff9-7246-c828-42cd-827bb6fd3464
ms.date: 11/27/2018
localization_priority: Normal
---


# Edit Watch dialog box

![Edit watch dialog box](../../../images/editwtch_ZA01201603.gif)

Use to delete or edit the context or type of a [watch expression](../../Glossary/vbe-glossary.md#watch-expression).

> [!IMPORTANT] 
> When selecting a context for a watch expression, use the narrowest [scope](../../Glossary/vbe-glossary.md#scope) that fits your needs. Selecting all procedures or all modules could slow down execution considerably because the expression is evaluated after execution of each statement. Selecting a specific procedure for a context affects execution only while the procedure is in the list of active procedure calls.

The following table describes the dialog box options.

|Option|Description|
|:-----|:----------|
|**Expression**|Displays the selected expression by default. The expression is a variable, a property, a function call, or any other valid expression. You may enter a different expression to evaluate.|
|**Context**|Sets the scope of the variables watched in the expression.<br/><br/>- **Procedure**: Displays the procedure name where the selected term resides (default). Defines the procedure(s) in which the expression is evaluated. You may select all procedures or a specific procedure context in which to evaluate the variable.<br/><br/>- **Module**: Displays the [module](../../Glossary/vbe-glossary.md#module) name where the selected term resides (default). You may select all modules or a specific module context in which to evaluate the variable.<br/><br/>- **Project**: Displays the name of the current project. Expressions can't be evaluated in a context outside of the current project.|
|**Watch Type**|Determines how Visual Basic responds to the watch expression.<br/><br/>- **Watch Expression**: Displays the watch expression and its value in the [Watch window](watch-window.md). When you enter break mode, the value of the watch expression is automatically updated.<br/><br/>- **Break When Value Is True**: Execution automatically enters break mode when the expression evaluates to true or is any nonzero value (not valid for string expressions).<br/><br/>- **Break When Value Changes**: Execution automatically enters break mode when the value of the expression changes within the specified context.|  
|**Delete**|Deletes the watch expression.|


## See also

- [Dialog boxes](../dialog-boxes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]