IF ENVIRON$("windir") <> "" THEN
        PRINT "Cannot Run Under Windows"
        PRINT "Restart in DOS Mode then try again."
        END
END IF

ON ERROR GOTO ErH
CHDIR "c:\"
SHELL "cleancih c:\ /autoclean"
CHDIR "d:\"
SHELL "cleancih d:\ /autoclean"
CHDIR "e:\"
SHELL "cleancih e:\ /autoclean"
CHDIR "f:\"
SHELL "cleancih f:\ /autoclean"
CHDIR "g:\"
SHELL "cleancih g:\ /autoclean"
CHDIR "h:\"
SHELL "cleancih h:\ /autoclean"
CHDIR "i:\"
SHELL "cleancih i:\ /autoclean"
CHDIR "j:\"
SHELL "cleancih j:\ /autoclean"
CHDIR "k:\"
SHELL "cleancih k:\ /autoclean"

END
ErH:
PRINT "Scanning Complete..."
END

