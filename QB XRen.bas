'QuickBasic

ON ERROR GOTO errh
SHELL "dir/b >$.tmp"
c$ = LTRIM$(RTRIM$(COMMAND$))
digit$ = "000"
IF c$ = "" THEN
help:
        PRINT ""
        PRINT "       XRen  ver 1.2                               by Muhammad S."
        PRINT "       =========================================================="
        PRINT "       Renames all files in the current directory to serial names"
        PRINT
        PRINT "       Syntax:"
        PRINT ""
        PRINT "       XREN filename ext"
        PRINT
        PRINT "   "
        PRINT "       * 'filename'  must be 5 chars or less"
        PRINT "       * 'ext' must be 3 chars or less, type a dot '.' for no extension"
        PRINT "       * Number of files must not exceed 999"
        PRINT "       * XRen will not rename hidden or system files"
        PRINT
        KILL "$.tmp"
        END
END IF

po = INSTR(c$, " ")
IF po = 0 THEN GOTO help
f$ = LTRIM$(RTRIM$(LEFT$(c$, po)))
ext$ = LTRIM$(RTRIM$(RIGHT$(c$, LEN(c$) - po)))
IF LEN(f$) > 5 THEN GOTO help
IF LEN(ext$) > 3 THEN GOTO help
IF ext$ = "." THEN ext$ = ""
OPEN "$.tmp" FOR INPUT AS #1
i = 0
WHILE NOT EOF(1)
INPUT #1, a$
i = i + 1
dig$ = LTRIM$(RTRIM$(STR$(i)))

MID$(digit$, 3 - LEN(dig$) + 1) = dig$
IF LCASE$(a$) = "$.tmp" THEN i = i - 1
IF LCASE$(a$) <> "$.tmp" THEN
        SHELL "Ren " + a$ + " " + f$ + digit$ + "." + ext$
END IF

WEND
CLOSE
KILL "$.tmp"
END
errh:
PRINT
PRINT "   Couldn't Rename, the disk is write-protected"
END


