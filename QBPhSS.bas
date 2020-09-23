x$ = LCASE$(COMMAND$)

IF x$ = "" OR x$ = "/?" OR x$ = "?" THEN
        PRINT
        PRINT "   PhtoSS        Screen Saver Generator           by MdSy"
        PRINT "   ======================================================="
        PRINT
        PRINT "       Syntax:    PHSS   bmpFileName"
        END
END IF

IF RIGHT$(x$, 4) = ".bmp" THEN x$ = LEFT$(x$, LEN(x$) - 4)
SHELL "copy /b sshdr.bin+" + x$ + ".bmp+" + "ssend.bin " + x$ + ".exe >nul"

