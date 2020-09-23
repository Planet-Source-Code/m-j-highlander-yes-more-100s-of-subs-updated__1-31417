DECLARE SUB CDROM ()
DIM a AS STRING
CLS
PLAY "l16abl32cdl64ef"
COLOR 15, 0

looper:
CLS
PRINT
PRINT
PRINT
PRINT
PRINT

PRINT "               Select an Option"
PRINT "               ****************"
PRINT
PRINT "               "; : COLOR 11, 0: PRINT "[S] SYSTEM:"; : COLOR 15, 0: PRINT " Transfer System Files to Hard Disk"
PRINT "               "; : COLOR 11, 0: PRINT "[F] FORMAT:"; : COLOR 15, 0: PRINT " Format Hard Disk"
PRINT "               "; : COLOR 11, 0: PRINT "[C] CD-ROM:"; : COLOR 15, 0: PRINT " Install CD-ROM Driver"
PRINT "               "; : COLOR 11, 0: PRINT "[X] Exit"
PRINT
a$ = INPUT$(1)
PLAY "l64ad"
IF ASC(a$) = 27 THEN END

IF a$ = "s" THEN a$ = "S"
IF a$ = "f" THEN a$ = "F"
IF a$ = "c" THEN a$ = "C"
IF a$ = "x" THEN a$ = "X"

SELECT CASE a$
        CASE "S"
        PRINT "Are you sure you want to copy System Files (y/n)";
        INPUT yn$
       
        IF yn$ = "y" OR yn$ = "Y" THEN
                SHELL "a:"
                PRINT "Copying System Files..."
                SHELL "sys c:"
                CDROM
        ELSE
                GOTO looper
        END IF

        CASE "F"
        SHELL "a:"
        PRINT "Are you sure you want to FORMAT Hard-Disk (y/n)";
                INPUT yn$
                IF yn$ = "y" OR yn$ = "Y" THEN
                SHELL "format c:/q/u/s/v:anwar"
                CDROM
        ELSE
                GOTO looper
        END IF

        CASE "C"
        SHELL "a:"
        CDROM
        
        CASE "X"
        PLAY "l64acac"
        END
       
        CASE ELSE
        GOTO looper
END SELECT

SUB CDROM
        PRINT "Installing CD-ROM Drivers..."
        SHELL "a:\cdsetup.exe"
        OPEN "c:\config.sys" FOR INPUT AS #1
        OPEN "c:\fuck.you" FOR OUTPUT AS #2
        PRINT #2, "LASTDRIVE=X"
        DO WHILE NOT EOF(1)
                LINE INPUT #1, L$
                PRINT #2, L$
        LOOP
        CLOSE #1, #2
        SHELL "a:\attrib -r -s -h c:\config.sys"
        KILL "c:\config.sys"
        NAME "c:\fuck.you" AS "c:\config.sys"
        PRINT
        PRINT
        PRINT "Remove Floppy Disk then press CTRL+ALT+DEL to reboot"

END SUB

