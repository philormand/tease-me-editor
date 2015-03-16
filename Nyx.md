![http://tease-me-editor.googlecode.com/svn/trunk/Wiki/Images/NyxTab.png](http://tease-me-editor.googlecode.com/svn/trunk/Wiki/Images/NyxTab.png)

# Description #
This will generate Nyx script that can be cut and pasted to the online Nyx editor.

You will need to upload the images, then cut and paste this generated script into the manual script window.

There are things that tease me can do that Nyx can't so not all teases will work.

I have created 2 flash teases (Kate's Birthday) and the Nyx script is just cut and pasted from here with no edits.

## Nyx stuff ##
Nyx only has the ability to do set / unset commands at the page level (not the delay / button level).  In Nyx you need to create a hidden page if you want to do this. The generator will automate this using numeric page names, starting with page 100 (or whatever you change it to at the top).

Nyx will only handle random pages if the page name is numeric (fred(1..2) will not go to fred1 or fred2) make sure you set the starting page number greater than any you have used to avoid it generating two pages with the same name.
From version 1.3 the editor will convert any random pages that use text to numbers.  I think the numbers will be unique but have not done a huge amount of testing.

(any suggestions on how to improve the Nyx generation or how to implement more tease me functions are welcome)

[Editor](http://code.google.com/p/tease-me-editor/wiki/WYSIWYG) [Instructions](http://code.google.com/p/tease-me-editor/wiki/Instructions)