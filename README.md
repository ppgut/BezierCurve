# BezierCurve / VBA 
Creation and manipulation of Bezier Curve on excel chart using chart events.

Chart events - Mouse Down, Mouse Up, Mouse Movement - are used to add functionalities to excel chart.

New functionalities:
Guiding points of Bezier Curve can be added by simple click on the chart. 
They can also be moved by click & drag.

Important thing in this project is the correct translation of mouse x, y position in pixels into points used by excel.
For this purpose 'GetDeviceCaps' function from gdi32 library is used to get screen x and y 'points per pixel' parameter.
Points can be then recalculated in a reference to chart position and axes size to determine X and Y values under mouse pointer.
