SCREEN 12
CLS
INPUT a$
INPUT x, y
CLS
PRINT " "; a$
l = LEN(a$) * 18
b = 20
FOR k = 1 TO y
	FOR i = 0 TO l
		FOR j = 0 TO 16
			IF POINT(i, j) = 0 THEN GOTO SKIP
			LINE (i * y + k, j * x - x + b)-(i * y + k, j * x + b), 3
SKIP:
		NEXT j
	NEXT i
NEXT k

