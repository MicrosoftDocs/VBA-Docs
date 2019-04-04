---
title: Application.FormulaBarHeight property (Excel)
keywords: vbaxl10.chm133306
f1_keywords:
- vbaxl10.chm133306
ms.prod: excel
api_name:
- Excel.Application.FormulaBarHeight
ms.assetid: ff377046-06cb-9cf7-32f5-773da447c184
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.FormulaBarHeight property (Excel)

Allows the user to specify the height of the formula bar in lines. Read/write **Long**.


## Syntax

_expression_.**FormulaBarHeight**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

If the specified value of **FormulaBarHeight** is greater than the viewable window space, the formula bar is resized to be equal to the window height.


## Example

In the following example, the height of the formula bar is set to five lines.


```vb
Application.FormulaBarHeight = 5 
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]