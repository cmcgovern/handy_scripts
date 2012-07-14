# -*- coding: utf-8 -*-

""" Excel Handy Functions """
__author__="Colin McGovern <colin@assembl.ie>"

# For any Excel column, returns integer equivalent (A = 0, B = 1, AA = 26, etc...)
# Converts all letters to upper case before processing
def convert_column_to_integer(column):
	"""Convert excel column letters into numbers that can be added/subtracted"""
	ordinal = 0
	position = 0
	
	# Convert string to letters and then reconstruct number (starting from the righ)
	for letter in reversed(list(column)):
		# Column can be A-Z, or AA, AB, AC, BA, BB, BC, etc...
		# This is base-26, hence conversion below and use of "position". 
		# Subtracts one from answer to start from zero, as with rows
		digit = ((ord(letter) - ord('A') + 1) * 26**position) - 1
		ordinal += digit
		position += 1
		
	return ordinal
	
# For any integer, returns the Excel equivalent
def convert_integer_to_column(number):
	"""Convert number equivalents back into Excel columns"""
	baseExcelNumber = baseExcel(number)
	
	# Correct initial letter, will be one too high (since 1 = "B", so e.g. 10 = "BA", not "AA")
	return decrement_letter(baseExcelNumber[0]) + baseExcelNumber[1:]
	
# Recursively build base 26 equivalent (sort of)
# There's probably a nicer way to do this using exclusively this function but
# I don't have time to work that out. All suggestions welcome
# The algorithm used here is based on really good answers here: 
# http://code.activestate.com/recipes/65212-convert-from-decimal-to-any-base-number/
def baseExcel(num, numerals="ABCDEFGHIJKLMNOPQRSTUVWYZ"):
	"""Convert from base10 to a sort of base 26 with changes (breaks for negative)"""
	return ((num == 0) and  "A" ) or (baseExcel(num // 26).lstrip("A") + numerals[num % 26])
	
# Return the letter before the current (or "A" if we can go no lower)
# Only expects caps
def decrement_letter(letter):
	"""Return previous letter (stops at 'A')"""
	return (letter == "A" and "A") or (chr(ord(letter) - 1))
	

