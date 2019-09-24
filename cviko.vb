sub test()

for i = 2 to range("A2").end(xldown).row
	cells(i,8) = cells(i,7)+cells(i,6)
next i

end sub
	
	
sub test()

	for i = 2 to range("B2").end(xldown).row
		cells(i,1) = "PL-" & cells(i,3) "-" cells(i,2)
	next i

end sub