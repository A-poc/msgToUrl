# msgToUrl
Python snippet for extracting URLs from outlook msg files.

```python
import re
import win32com.client
import os
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
for name in os.listdir("."):
    if name.endswith(".msg"):
        msg = outlook.OpenSharedItem((os.path.dirname(os.path.realpath(__file__)))+"\\"+name)
        match = re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', msg.Body)
        for m in match:
            print(m.replace(">",""))
```

**Requirements**
- Outlook 'running'
- Windows
- pywin32 (*pip install pywin32*)
- .msg files in script folder
