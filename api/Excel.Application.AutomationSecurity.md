---
title: Application.AutomationSecurity property (Excel)
keywords: vbaxl10.chm133269
f1_keywords:
- vbaxl10.chm133269
ms.prod: excel
api_name:
- Excel.Application.AutomationSecurity
ms.assetid: ae19bf93-dc0f-f18a-d8ce-f54108602844
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.AutomationSecurity property (Excel)

Returns or sets an **[MsoAutomationSecurity](Office.MsoAutomationSecurity.md)** constant that represents the security mode that Microsoft Excel uses when programmatically opening files. Read/write.


## Syntax

_expression_.**AutomationSecurity**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

This property is automatically set to **msoAutomationSecurityLow** when the application is started. Therefore, to avoid breaking solutions that rely on the default setting, you should be careful to reset this property to **msoAutomationSecurityLow** after programmatically opening a file. Also, this property should be set immediately before and after opening a file programmatically to avoid malicious subversion.

**MsoAutomationSecurity** can be one of these **MsoAutomationSecurity** constants:

- **msoAutomationSecurityByUI**. Uses the security setting specified in the **Security** dialog box.

- **msoAutomationSecurityForceDisable**. Disables all macros in all files opened programmatically without showing any security alerts. 

  > [!NOTE] 
  > This setting does not disable Microsoft Excel 4.0 macros. If a file that contains Microsoft Excel 4.0 macros is opened programmatically, the user will be prompted to decide whether to open the file.

- **msoAutomationSecurityLow**. Enables all macros. This is the default value when the application is started.

Setting **[ScreenUpdating](Excel.Application.ScreenUpdating.md)** to **False** does not affect alerts and will not affect security warnings. 

The **[DisplayAlerts](Excel.Application.DisplayAlerts.md)** setting will not apply to security warnings. For example, if the user sets **DisplayAlerts** equal to **False** and **AutomationSecurity** to **msoAutomationSecurityByUI** while the user is on Medium security level, there will be security warnings while the macro is running. This allows the macro to trap file open errors, while still showing the security warning if the file open succeeds.


## Example

This example captures the current automation security setting, changes the setting to disable macros, displays the **Open** dialog box, and after opening the selected document, sets the automation security back to its original setting.

```vb
Sub Security() 
    Dim secAutomation As MsoAutomationSecurity 
 
    secAutomation = Application.AutomationSecurity 
 
    Application.AutomationSecurity = msoAutomationSecurityForceDisable 
    Application.FileDialog(msoFileDialogOpen).Show 
 
    Application.AutomationSecurity = secAutomation 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
