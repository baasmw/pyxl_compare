## Collect data from separate Excel files and output the cities that exist in both files
import os, sys, openpyxl
original_cities = []
compared_cities = []
difference = []

## Location of script
location = os.path.realpath( os.path.join( os.getcwd(), os.path.dirname( sys.argv[0] ) ) )

## Read out xlsx files
file1 = openpyxl.load_workbook( os.path.join( location, 'original.xlsx' ) )
file2 = openpyxl.load_workbook( os.path.join( location, 'reference.xlsx' ) )

## Get content from the first sheet of both files
original = file1.get_sheet_by_name( file1.sheetnames[0] )
reference = file2.get_sheet_by_name( file2.sheetnames[0] )

## Extract content from selected sheets
def get_cities( citylist, output ):
	for city in citylist['A1':'A' + str( citylist.max_row )]:
		for name in city:
			output.append( name.value )

get_cities( original, original_cities )
get_cities( reference, compared_cities )

## Collect results in an array
for value in original_cities:
	if value in compared_cities:
		difference.append( value )

## Output the matching values
print( str( len( difference ) ) + " matches:" )
for result in difference:
	print( result )