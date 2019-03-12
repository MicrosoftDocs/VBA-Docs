---
title: Compiler constants (VBA)
keywords: vblr6.chm1020792
f1_keywords:
- vblr6.chm1020792
ms.prod: office
ms.assetid: bde15ce4-af30-1bbf-7d34-4cfa7e396261
ms.date: 12/21/2018
localization_priority: Normal
---


# Compiler constants

Visual Basic for Applications defines [constants](../../Glossary/vbe-glossary.md#constant) for exclusive use with the **[#If...Then...#Else](../../reference/user-interface-help/ifthenelse-directive.md)** directive. These constants are functionally equivalent to constants defined with the **#If...Then...#Else** directive except that they are global in [scope](../../Glossary/vbe-glossary.md#scope); that is, they apply everywhere in a [project](../../Glossary/vbe-glossary.md#project).

> [!NOTE] 
> Because **Win32** returns **True** in both 32-bit and 64-bit development platforms, it is important that the order within the **#If...Then...#Else** directive returns the desired results in your code. For example, because **Win32** returns **True** in 64-bit (**Win32** is compatible in **Win64** environments), checking for **Win32** before **Win64** results in the **Win64** condition never running because **Win32** returns **True**. The following order returns predictable results (this applies to both Winx and VBAx constants):
> 
>  ```vb
>  #If Win64 Then 
>  ' Win64=true, Win32=true, Win16= false 
>  #ElseIf Win32 Then 
>  ' Win32=true, Win16=false 
>  #Else 
>  ' Win16=true 
>  #End If
>  ```

<br/>

On 16-bit development platforms, the compiler constants are defined as follows.

|Constant|Value|Indicates that the development environment... |
|:-----|:-----|:-----|
|**Win16**|**True**|Is 16-bit compatible.|
|**Win32**|**False**|Is not 32-bit compatible.|
|**Win64**|**False**|Is not 64-bit compatible.|

<br/>

On 32-bit development platforms, the compiler constants are defined as follows.

|Constant|Value|Indicates that the development environment...|
|:-----|:-----|:-----|
|**Vba6**|**True**|Is Visual Basic for Applications, version 6.0 compatible.|
|**Vba6**|**False**|Is not Visual Basic for Applications, version 6.0 compatible.|
|**Vba7**|**True**|Is Visual Basic for Applications, version 7.0 compatible.|
|**Vba7**|**False**|Is not Visual Basic for Applications, version 7.0 compatible.|
|**Win16**|**False**|Is not 16-bit compatible.|
|**Win32**|**True**|Is 32-bit compatible.|
|**Win64**|**False**|Is not 64-bit compatible.|
|**Mac**|**True**|Is Macintosh.|
|**Mac**|**False**|Is not Macintosh.|


<br/>

On 64-bit development platforms, the compiler constants are defined as follows.

|Constant|Value|Indicates that the development environment...|
|:-----|:-----|:-----|
|**Vba6**|**True**|Is Visual Basic for Applications, version 6.0 compatible.|
|**Vba6**|**False**|Is not Visual Basic for Applications, version 6.0 compatible.|
|**Vba7**|**True**|Is Visual Basic for Applications, version 7.0 compatible.|
|**Vba7**|**False**|Is not Visual Basic for Applications, version 7.0 compatible.|
|**Win16**|**False**|Is not 16-bit compatible.|
|**Win32**|**True**|Is 32-bit compatible.|
|**Win64**|**True**|Is 64-bit compatible.|
|**Mac**|**True**|Is Macintosh.|
|**Mac**|**False**|Is not Macintosh.|

> [!NOTE] 
> These constants are provided by Visual Basic, so you cannot define your own constants with these same names at any level.


## See also

- [Understanding conditional compilation](understanding-conditional-compilation.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
