Using Python to Control LARSA 4D
================================

LARSA 4D macros can be written using the pywin32 package. See http://timgolden.me.uk/pywin32-docs/html/com/win32com/HTML/QuickStartClientCom.html for more information about that package. Last tested with Python 3.9.7.

Because there is no autocomplete in Python editors and no Object Browser as in the Excel macro editor, writing macros with Python is not a very friendly process.

To install the pywin32 package, in a Command Prompt run:

    py -3 -m pip install pywin32
    
LARSA 4D Application Object
---------------------------

The Application object can be used for cross-process communication with a running instance of LARSA 4D.

```py
import win32com.client
larsa4d = win32com.client.Dispatch("LARSA2000.Application")
larsa4d.showLARSA()
```
