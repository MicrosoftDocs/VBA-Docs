---
title: NavigationControl.KeyboardLanguage property (Access)
keywords: vbaac10.chm11131
f1_keywords:
- vbaac10.chm11131
ms.prod: access
api_name:
- Access.NavigationControl.KeyboardLanguage
ms.assetid: 5a4f4c8b-2d01-4613-2bb0-8c3e2c7dfda9
ms.date: 03/02/2019
localization_priority: Normal
---


# NavigationControl.KeyboardLanguage property (Access)

## Syntax

_expression_.**KeyboardLanguage**

_expression_ A variable that represents a **[NavigationControl](Access.NavigationControl.md)** object.


## Remarks

Valid values for this property are 0 (zero), which corresponds to the default system language, or _plid_ + 2, where _plid_ is the primary language ID of a language installed on the current system. For example, the primary language ID of English is 9, so the corresponding **KeyboardLanguage** setting is 11. 

For a list of languages and their primary language IDs, see [Language Identifier Constants and Strings](https://docs.microsoft.com/windows/desktop/Intl/language-identifier-constants-and-strings). An exception to this list is Traditional Chinese, which is represented by the value 200.

Setting this property to a language that is not installed may either have no effect or cause an error.

## See also

- [Language Identifiers](https://docs.microsoft.com/windows/desktop/intl/language-identifiers)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]