---
title: Application.SendKeys method (Excel)
keywords: vbaxl10.chm183108
f1_keywords:
- vbaxl10.chm183108
ms.prod: excel
api_name:
- Excel.Application.SendKeys
ms.assetid: 585666b9-adbc-5d04-c480-58e55ea7fb9d
ms.date: 04/05/2019
localization_priority: Priority
---


# Application.SendKeys method (Excel)

Sends keystrokes to the active application.


## Syntax

_expression_.**SendKeys** (_Keys_, _Wait_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Keys_|Required| **Variant**|The key or key combination that you want to send to the application, as text.|
| _Wait_|Optional| **Variant**| **True** to have Microsoft Excel wait for the keys to be processed before returning control to the macro. **False** (or omitted) to continue running the macro without waiting for the keys to be processed.|

## Remarks

This method places keystrokes in a key buffer. In some cases, you must call this method before you call the method that will use the keystrokes. For example, to send a password to a dialog box, you must call the **SendKeys** method before you display the dialog box.

The _Keys_ argument can specify any single key or any key combined with Alt, Ctrl, or Shift (or any combination of those keys). Each key is represented by one or more characters, such as `"a"` for the character a, or `"{ENTER}"` for the Enter key.

To specify characters that aren't displayed when you press the corresponding key (for example, Enter or Tab), use the codes listed in the following table. Each code in the table represents one key on the keyboard.

|Key|Code|
|:-----|:-----|
|BACKSPACE| `{BACKSPACE}` or `{BS}`|
|BREAK| `{BREAK}`|
|CAPS LOCK| `{CAPSLOCK}`|
|CLEAR| `{CLEAR}`|
|DELETE or DEL| `{DELETE}` or `{DEL}`|
|DOWN ARROW| `{DOWN}`|
|END| `{END}`|
|ENTER (numeric keypad)| `{ENTER}`|
|ENTER| `~` (tilde)|
|ESC| `{ESCAPE}` or `{ESC}`|
|HELP| `{HELP}`|
|HOME| `{HOME}`|
|INS| `{INSERT}`|
|LEFT ARROW| `{LEFT}`|
|NUM LOCK| `{NUMLOCK}`|
|PAGE DOWN| `{PGDN}`|
|PAGE UP| `{PGUP}`|
|RETURN| `{RETURN}`|
|RIGHT ARROW| `{RIGHT}`|
|SCROLL LOCK| `{SCROLLLOCK}`|
|TAB| `{TAB}`|
|UP ARROW| `{UP}`|
|F1 through F15| `{F1}` through `{F15}`|

<br/>

You can also specify keys combined with Shift and/or Ctrl and/or Alt. To specify a key combined with another key or keys, use the following table.

|To combine a key with|Precede the key code with|
|:-----|:-----|
|Shift| `+` (plus sign)|
|Ctrl| `^` (caret)|
|Alt| `%` (percent sign)|

## Example

The following example creates a new workbook.

```vb
Application.SendKeys("^n")
```

The following example displays the Name Manager.

```vb
Application.SendKeys("%mn")
```

The following example enters the value 1234 into the Active Cell.

```vb
Application.SendKeys ("1234{Enter}")
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
