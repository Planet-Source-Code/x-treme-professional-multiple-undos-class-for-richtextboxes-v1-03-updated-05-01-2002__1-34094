I tried a lot of differnet multi undo examples from the Inet, but every had at
least one draw back. So I had to write my on solution and I think this one is
quite smart.
 
If you want to use it, you don't have to clutter your code by pasting in lots
of undo-related subs, functions and variables.

Everything comes within two handy class files.
 
All you have to do is to add both classes to your VB project, create one class
instance of clsUndo and select the richtextbox you want to enhance with
multiple undos.

The code you have to add will look something like this:
 
  Dim WithEvents Undo as clsUndo
  Set Undo = New clsUndo
  Undo.Create Controls, RichTextBox1
 
The create method of clsUndos needs two arguments. 
The first one is a Controls object, where the class can create a timer object
on runtime. The second argument is your richtextbox.
An optional third parameter is the undo delay time.


New in version 1.03:
--------------------
Speed optimization by changing the algorithm in getMatchingCharCount.
New properties: getUndoCount, getRedoCount
Undo and Redo Method can have an argument where you can indicate the count steps you want to undo/redo
 
New in version 1.02:
--------------------
You can turn on/off the automatical tracking of text changes.
This may be useful on automated text operations.
The cut/copy/paste buttons now support pictures and other embedded elements.

New in version 1.01:
--------------------
The single undo items do not store the whole text anymore. Instead they only
contain the changed text.
This is important for editing bigger files. I.e. when you are editing a 500kb
File than you would need 500kB for every undo item.
That way you can easily run out of memory.


If you like this code, please vote for it at Planet-Source-Code.com:
http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=34094


Kind regards,
Sebastian Thomschke										  			05/01/2002
http://www.sebthom.de