# kex
STL export for AutoCAD and import script for Cinema4D

- this script explode all blocks and export all solids as .stl to a folder.

- all informations about the exploded blocks and solids are stored in a .xml file.

- it's also a kex.py/kex.tif file included for cinema4d. this script can read the .xml file and import all exported .stl files and group them like the original blocks in autocad.
copy this 2 files to your cinema4d library/scripts folder.

- the build folder contain the ready to use .dll and a autoload .lsp file for autocad
- place the .dll and the .lsp file to a thrusted location (autocad options).
- use the APPLOAD command and add the autoload.lsp into the content aera.
- type KEX to start the pallete.

- if you place the files on a different location than c:\vbnet remember to change the path in the .lsp file and in the KEX pallete options tab (export location).

# KEX palette options (Export tab)
- export solids that are not in a block -> means all solids in the drawing (without blocks)
- export hidden, frozen or off objects in exploded blocks -> means if a solid is on a off layer (and invisible) export them. the same with frozen or hidden.
- don't export blocks on specific layers -> means if the blocklayer is on one of this layers, this block is ignored during the export. use , (comma) to seperate multiple layers.
- FILEDIA = 1 button -> the systemvariable FILEDIA is changed during the export (FILEDIA 0) if something go wrong during export, you can restore this variable with this button (or type FILEDIA and set it to 1)

# KEX palette options (Options tab)
- export path -> the location for the .stl files. it will create a sub-folder with the current username.
- quality -> the mesh quality, 0.01 = bad, 10 = best
- reset ucs -> reset the coordinate system (important for moving objects)
- move all objects + 10'000 in xyz -> make sure all objects have + coordinates!
- close the new drawing after export -> the export don't change your original drawing. it will create a temporary drawing and do all the work there. uncheck this option if you want see the result with the exploded blocks and new layers.
- prefix for temp. object layer in blocks -> identificator that the object layer was a block.
- custom properties - C4D import root name -> from your drawing (custom properties) get the value from the selected key. This value is used to rename the root object with the cinema4d import script.
