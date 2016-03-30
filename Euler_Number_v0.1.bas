REM  *****  BASIC  *****

Sub Main

REM Calculate Euler's Number
REM
REM author: D. Barland date: March 22, 2016
REM Version: 0.1
REM
REM Purpose: Calculate Euler's number using the
REM formula e = Sum{0...n} (1 + 1/n!)

	Dim euler, e, abs_error As Double
	Dim i, n, n_max, n_factorial As Long

	euler = exp(1)
	'print "euler e=", euler

	' set n_max:{3,5,7,10,12}
	n_max = 10

	e = 1 'sets starting value of e to the first term of 1/0!

	'loop to calculate Euler's number:	
	For n = 1 to n_max

		'compute factorial of n
		n_factorial = 1
		for i = 1 to n
			n_factorial = n_factorial * i
		next
		
		'compute e
		e = e + 1 / n_factorial

	Next

	'compute absolute error
	abs_error = abs(euler - e)

	'print "for n=", n," euler's number e=", e, " the error is ", abs_error

	'===================================================================
	'The following section writes our calculations to spreadsheet cells
	' below block must be commented out as it will not work for MS Office
	dim document   as object
	dim dispatcher as object
 
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
 
	dim args1(0) as new com.sun.star.beans.PropertyValue
	dim args2(0) as new com.sun.star.beans.PropertyValue
	dim n_cell, e_cell, err_cell as String 	
	'	choose cells in which to write labels and calculation results
	n_label_cell = "$A$2"
	n_cell = "$A$3"
	e_label_cell = "$C$2"
	e_cell = "$C$3"
	err_label_cell = "$E$2"
	err_cell = "$E$3"

	'	transfer n label to spreadsheet
	args1(0).Name = "ToPoint"
	args1(0).Value = n_label_cell
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
 
	args2(0).Name = "StringName"
	args2(0).Value = "n"
	dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args2())

	'	transfer e label to spreadsheet
	args1(0).Name = "ToPoint"
	args1(0).Value = e_label_cell
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
 
	args2(0).Name = "StringName"
	args2(0).Value = "e"
	dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args2())

	'	transfer error label to spreadsheet
	args1(0).Name = "ToPoint"
	args1(0).Value = err_label_cell
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
 
	args2(0).Name = "StringName"
	args2(0).Value = "error"
	dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args2())

	'	transfer n value to spreadsheet
	args1(0).Name = "ToPoint"
	args1(0).Value = n_cell
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
 
	args2(0).Name = "StringName"
	args2(0).Value = Str(n_max)
	dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args2())

	'	transfer e value to spreadsheet
	args1(0).Name = "ToPoint"
	args1(0).Value = e_cell
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
 
	args2(0).Name = "StringName"
	args2(0).Value = Str(e)
	dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args2())

	'	transfer error value to spreadsheet
	args1(0).Name = "ToPoint"
	args1(0).Value = err_cell
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
 
	args2(0).Name = "StringName"
	args2(0).Value = Str(abs_error)
	dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args2())
	'===================================================================

End Sub

