---
title: MacScript function (Visual Basic for Applications)
keywords: vblr6.chm1010848
f1_keywords:
- vblr6.chm1010848
ms.prod: office
ms.assetid: d845de85-a0d8-e10e-1174-8571e42bb8d2
ms.date: 12/13/2018
localization_priority: Normal
---

# MacScript function

> [!NOTE] 
> This function has been deprecated, therefore it is no longer supported. For more information, see this [Stack Overflow article](https://stackoverflow.com/a/30949324/209942).

Executes an AppleScript script and returns a value returned by the script, if any.

## Syntax

**MacScript**(_script_)

The _script_ argument is a [String expression](../../Glossary/vbe-glossary.md#string-expression). The **String** expression either can be a series of AppleScript commands or can specify the name of an AppleScript script or a script file.

## Remarks

Multiline scripts can be created by embedding carriage-return characters (**Chr**(13)).

```vb
ThePath$ = Macscript("ChooseFile")

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]