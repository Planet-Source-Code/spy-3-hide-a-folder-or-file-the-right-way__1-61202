<div align="center">

## Hide a folder or file the right way\.

<img src="PIC2005617125385391.JPG">
</div>

### Description

This code will hide any folder or file from windows explorer(By hide I mean even from those who have View hidden files and folders checked on View Screen Shot)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SPY\-3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/spy-3.md)
**Level**          |Beginner
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/spy-3-hide-a-folder-or-file-the-right-way__1-61202/archive/master.zip)





### Source Code

I was just informed about a built in vb function SetAttr %File or Folder%, vbSystem or vbHidden or vbNormal<br>My old code is below still.
I was messing around with some code and found how to hide folders and files(or atleast turn it into a system folder or file).<br>(if you were making a secure folder why not hide it the right way so you cant see it even if you have show hidden files and folders checked?)<br> The actual code is very very simple that anyone could do it (yet for some reason they dont) Heres the code:<br>
To hide a folder simple use this code<br><font color="blue"><b>
Dim FS, F<br>
Set FS = CreateObject("Scripting.FileSystemObject")<br>
Set F = FS.GetFolder(%FOLDERPATH%) <font color="green">'Replace %FOLDERPATH% with the folders path</font><br>
F.Attributes = -1 <font color="green">' -1 Makes it a system folder so its hidden from windows explorer(works for me on xp)</font><br>
</font></b>To unhide the folder simply put<br><font color="blue"><b>
Dim FS, F<br>
Set FS = CreateObject("Scripting.FileSystemObject")<br>
Set F = FS.GetFolder(%FOLDERPATH%) <font color="green">'Replace %FOLDERPATH% with the folders path </font><br>
F.Attributes = 0 <font color="green">' This returns the folder to normal in windows explorer</font>
</font></b><br>To hide files simply use this code<br><font color="blue"><b>
Dim FS, F<br>
Set FS = CreateObject("Scripting.FileSystemObject")<br>
Set F = FS.GetFile(%FILEPATH%) <font color="green">'Replace %FILEPATH% with the files path</font><br>
F.Attributes = -1 <font color="green">' -1 Makes it a system file so its hidden from windows explorer(works for me on xp)</font></b><br></font>
To unhide the file simply put<br><font color="blue"><b>
Dim FS, F<br>
Set FS = CreateObject("Scripting.FileSystemObject")<br>
Set F = FS.GetFile(%FILEPATH%) <font color="green">'Replace %FILEPATH% with the files path</font><br>
F.Attributes = 0 <font color="green">' This returns the file to normal in windows explorer</font>
</font></b>
<br><b>
<br>Hope this code helps and please leave feedback and/or vote.</b>

