SCREEN 12
CLS
pi = 4 * ATN(1)
a = 99
b = 11

g = (a + b) / b
h = (b - a) / b

FOR t = 0 TO 2 * pi STEP .01
        ' a = a + .1  ''try this!
        ' a = a + .1: b = b + .1  ''also this!
        x = (a + b) * COS(t) - b * COS(g * t) + 300
        y = (a + b) * SIN(t) - b * SIN(g * t) + 200

        'x = (a - b) * COS(t) + b * COS(h * t) + 300
        'y = (a - b) * SIN(t) + b * SIN(h * t) + 200

        PSET (x, y)
NEXT t

