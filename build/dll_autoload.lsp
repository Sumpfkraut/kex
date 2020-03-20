(if (findfile "C:\\vbnet\\kex.dll")
  (progn
    (princ "C:\\vbnet\\kex.dll loaded.\n")
    (command "._NETLOAD" "C:\\vbnet\\kex.dll"))
  (princ "C:\\vbnet\\kex.dll NOT found!\n"))