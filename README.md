# kex
STL export for AutoCAD

- this script explode all blocks and export all solids as .stl to a folder.

- all informations about the exploded blocks and solids are stored in a .xml file.

- it's also a kex.py/kex.tif file included for cinema4d. this script can read the .xml file and import all exported .stl files and group them like the original blocks in autocad.
copy this 2 files to your cinema4d library/scripts folder.

- the build folder contain the ready to use .dll and a autoload .lsp file for autocad
- place the .dll and the .lsp file to a thrusted location (autocad options).
- use the APPLOAD command and add the autoload.lsp into the content aera.
