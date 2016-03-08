REM  *****  BASIC  *****

Sub Main

REM – Exponential Equation Solution by Iteration
REM
REM – author: D. Barland date: February 23, 2016
REM – Version: 0.1
REM
REM – Purpose: Calculate exponential equation solution by iteration
REM - Exponential equation is in the form: y = a times b to the x power

	Dim a, b, x, y, y_guess
	Dim precision
	Dim I

	a = 1
	b = 5
	y = .4

	precision = 5	'desired decimal precision of solution

REM - First step, find two integers between which 'the solution' lies
REM - We will restrict ourselves to a domain {-10 <= x <= 10}
REM - x_lower is the less of the two integers
REM - x_upper is the greater of the two integers

	Dim x_lower, x_upper
	
	'start x_lower at the bottom of our domain:
	x_lower = -10
	'then x_upper is one plus x_lower
	x_upper = -9
	
	'loop to find the two 'bounding' integers:
	do
		Dim still_getting_bounds
		still_getting_bounds = true
		'calculate the two corresponding exponential equation values
		y_lower = a * b^(x_lower)
		y_upper = a * b^(x_upper)
		
		'now test to see if the desired function value lies between y_lower and y_upper
		if (y >= y_lower) and (y <= y_upper) then
			'if solution is bounded, we have our two integers!
			print x_lower, x_upper
			still_getting_bounds = false
		else
			'if solution does not lie between the current integers, we increase them
			x_lower = x_lower + 1
			x_upper = x_lower + 1 'Can you see why this works?
		end if
	loop while still_getting_bounds

REM - Second step, 'iterate' the variable 'x' to converge to the solution
REM - 'epsilon' is our convergence tolerance, this is how we know if we are satisfied

	Dim epsilon
	epsilon = .1^precision
	
	'lets choose our starting 'x' to be halfway between our x_lower and x_upper
	x = (x_lower + x_upper) / 2
	
	do while abs(y - y_guess) > epsilon
		y_guess = a * b^(x)
		if abs(y - y_guess) > epsilon then
			if y > y_guess then
				x = x + epsilon/2	'use one half epsilon to get closer to solution
			elseif y < y_guess then
				x = x - epsilon/2
			endif
		elseif abs(y - y_guess) < epsilon then
			'We have our solution!
			print "solution x = ", x
		endif
	loop

	'now check our answer
	y_final = a*b^x
	print "our final y = ", y_final
End Sub
