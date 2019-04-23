---
title: Form.DatasheetGridlinesColor property (Access)
keywords: vbaac10.chm13403
f1_keywords:
- vbaac10.chm13403
ms.prod: access
api_name:
- Access.Form.DatasheetGridlinesColor
ms.assetid: 92d07c1c-fc47-0049-7da3-a34ee56fbc83
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.DatasheetGridlinesColor property (Access)

You can use the **DatasheetGridlinesColor** property to specify the color of gridlines in a datasheet. Read/write **Long**.


## Syntax

_expression_.**DatasheetGridlinesColor**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **DatasheetGridlinesColor** property applies only to objects in Datasheet view.

This property is only available in [Visual Basic](../access/Concepts/Settings/set-properties-by-using-visual-basic.md) within a Microsoft Access database.

You can also use the **RGB** or **QBColor** functions to set this property.

This property setting affects the gridline color for the entire datasheet. It's not possible to set the gridline color of individual cells in Datasheet view.

The following table contains the properties that don't exist in the DAO **Properties** collection until you set them by using the **Formatting (Datasheet)** toolbar, or you can add them in an Access database by using the **CreateProperty** method and append it to the DAO **Properties** collection.

|||
|:-----|:-----|
|**[DatasheetFontItalic](Access.Form.DatasheetFontItalic.md)** *|**[DatasheetForeColor](Access.Form.DatasheetForeColor.md)** *|
|**[DatasheetFontHeight](Access.Form.DatasheetFontHeight.md)** *|**[DatasheetBackColor](Access.Form.DatasheetBackColor.md)**|
|**[DatasheetFontName](Access.Form.DatasheetFontName.md)** *|**DatasheetGridlinesColor**|
|**[DatasheetFontUnderline](Access.Form.DatasheetFontUnderline.md)** *|**[DatasheetGridlinesBehavior](Access.Form.DatasheetGridlinesBehavior.md)**|
|**[DatasheetFontWeight](Access.Form.DatasheetFontWeight.md)** *|**[DatasheetCellsEffect](Access.Form.DatasheetCellsEffect.md)**|

> [!NOTE] 
> When you add or set any property listed with an asterisk, Access automatically adds all the properties listed with an asterisk to the **Properties** collection in the database.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]