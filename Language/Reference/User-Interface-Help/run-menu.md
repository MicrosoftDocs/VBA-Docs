---
title: Run menu
keywords: vbui6.chm2057562
f1_keywords:
- vbui6.chm2057562
ms.prod: office
ms.assetid: 6a60dc31-5a3d-b72b-40ea-309ec6a1e044
ms.date: 11/21/2018 
---


# Run menu

|Command|Description|Toolbar button|Keyboard shortcut|
|:------|:----------|:-------------|:----------------|
|**Break** | Stops execution of a program while it's running and switches to [break mode](../../Glossary/vbe-glossary.md#break-mode).<br/><br/>Any statement being executed when you choose this command is displayed in the Code window with ![Breakpoint](../../../images/wbrkpnt_ZA01201808.gif) in the left margin if you selected the Margin Indicator Bar on the **Editor Format** tab of the **Options** dialog box.<br/><br/>If the application is waiting for events in the idle loop (no statement is being executed), no statement is highlighted until an event occurs.<br/><br/>Some editing changes made in break mode may require you to restart your program for the changes to take effect.|![Toolbar button](../../../images/tbr_brk_ZA01201682.gif) |CTRL+BREAK |
|**Design Mode** |Turns design mode on per project and then changes to **Exit Design Mode**. Design mode is the time during which no code from the project is running and events from the host or project will not execute. You can leave design mode by executing a macro or using the Immediate window.| | |
|**Exit Design Mode** |Turns design mode off per project and clears all module level variables in the project.|![Toolbar button](../../../images/tbr_dsgm_ZA01201699.gif)| |
|**Reset Project** | Clears the **Call** stack and clears the module level variables.|![Toolbar button](../../../images/tbr_end_ZA01201701.gif)| |
|**Run Project**| Puts the project into a mode in which it can be used by other applications. This is used to debug and test the stand-alone project before building a [dynamic-link library (DLL)](../../Glossary/vbe-glossary.md#dynamic-link-library-dll) (DLL) from it. The current project is registered, replacing any existing registration information for the project (the registry information for an existing DLL version of the project, for example).| | |
|**Stop Project**|Unregisters the project, and restores any previous registry information. This makes the in-memory project no longer able to be called from other applications.<br/><br/>**NOTES**: The **Run Project** and **Stop Project** commands are available only to the current stand-alone project. They are not available to [host application](../../Glossary/vbe-glossary.md#host-application) document projects. This feature is not available in all versions of the Visual Basic Editor.| | |
|**Run Sub/UserForm**|Runs the current procedure if the cursor is in a procedure, or runs the form if a form is currently active. This command becomes the **Continue** command when you are in [break mode](../../Glossary/vbe-glossary.md#break-mode).<br/><br/>If neither the [Code window](code-window.md) nor the **[UserForm](userform-window.md)** window is active, this command becomes the **Run Macro** command.| | |
|**Continue**|Resumes running the current procedure or form.| | |
|**Run Macro**|Runs the macro.| | |