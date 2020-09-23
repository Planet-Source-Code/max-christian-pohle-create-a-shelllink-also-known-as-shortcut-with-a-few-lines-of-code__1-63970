<div align="center">

## create a Shelllink \(also known as Shortcut\) with a few lines of code


</div>

### Description

It creates a shortcut and lets you define all important parameters. I have seen some examples before, but they where all too complicated for this little job. So I collected informationen and wrote this.
 
### More Info
 
for example

CreateShortcut App.Path &amp; "\Explorer.lnk", "C:\winnt\explorer.exe", "", "CTRL+SHIFT+D", "Microsoft Windows Explorer", Maximized, 0

creates a link in your applications path which calls the explorer.exe from "c:\winnt", defines CTRL+Shift+D as keycombination for it, starts the Explorer-window maximized and uses the standardicon of explorer.exe, which is 0.

Returns? The shortcut :-)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Max Christian Pohle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/max-christian-pohle.md)
**Level**          |Beginner
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/max-christian-pohle-create-a-shelllink-also-known-as-shortcut-with-a-few-lines-of-code__1-63970/archive/master.zip)

### API Declarations

```
no APIs used, but Windows-Scriptinghost
Thats why it should run as "VB Script" as well ( with little modifications maybe :-) )
```


### Source Code

```
Option Explicit
Enum Windowstyle
  Minimized = 7
  Maximized = 3
  Normal = 1
End Enum
Public Function CreateShortcut(Linkpath As String, TargetPath As String, Optional WorkPath As String, Optional HotKey As String = "", Optional Description As String = "", Optional Winstyle As Windowstyle, Optional Iconnumber As Integer)
  Dim SC As Object
  Set SC = CreateObject("Wscript.Shell").CreateShortcut(Linkpath)
  With SC
    .TargetPath = TargetPath
    'where your shortcuts jumps to
    .HotKey = HotKey
    'can be "CTRL+SHIFT+E" (as Str!) for Example
    .Description = Description
    'this should be clear to you
    .Windowstyle = Winstyle
    'Winstyle differs from the typical styles (2 does not mean maximized)
    .IconLocation = TargetPath & ", " & Iconnumber
    'This will take the Icon for the link from the file its associated with (targetpath)
    'some files include more than one icon. This is what is meant by Iconnumber.
    .WorkingDirectory = WorkPath
    'The normal Workingdirectory for the file the link calls.
    .Save
    'saves the link [important! :-)]
  End With
End Function
```

