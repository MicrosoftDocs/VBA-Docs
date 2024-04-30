---
title: Auto Macros (PowerPoint VBA reference)
ms.date: 04/28/2024
ms.localizationpriority: medium
---

# Auto Macros (PowerPoint VBA reference)

By giving a macro a special name, you can run it automatically when you perform an operation such as starting PowerPoint. 
PowerPoint recognizes the following names as automatic macros, or "auto" macros.

|**Macro name**|**Run conditions**|
|:-----|:-----|
|Auto_Open|Runs when PowerPoint is started or it has loaded the add-in|
|Auto_Close|Runs when you exit PowerPoint or unloads the add-in|

## PowerPoint auto macros behavior

The behavior of auto macros `Auto_Open` and `Auto_Close` are different from a similar macros `AutoOpen` and `AutoClose` in Word and Excel.

PowerPoint will run the `Auto_Open` each time it loads an add-in (`*.ppam`) with a module containing this macro.
It will effectively run when PowerPoint application is started or you add the add-in
using the **PowerPoint Add-ins** dialog.

PowerPoint will run the `Auto_Close` each time it unloads an add-in with a module containing this macro.
It will effectively run when PowerPoint application is being shutdown or your unload the add-in
using the **PowerPoint Add-ins** dialog.
