Sub MALFUNCTION()

Set asheet = Worksheets("MALFUNCTION")

'Clear Existing Data

asheet.Range("E3:E86") = "" 

Dim i, j, k

'Search All The MALFUNCTION 

For i = 2 To 50

	If asheet.Cells(i, 7) <> "" Then

	'MALFUNCTION Customer name 

	na = asheet.Cells(i, 7)

	'MALFUNCTION Duration

	du = asheet.Cells(i, 11)

	'MALFUNCTION Column

	co = asheet.Cells(i, 12)

	'MALFUNCTION Shutdown type

	st = asheet.Cells(i, 18)

	'Match The Customer name By turns

		For j = 1 To 7

			If na = asheet.Cells(j + 1, 20) Then

				'QEI,TEI,NEI

				If st = "External Interruptions" Or st = "Disuse by Customer" Then

				asheet.Cells(2 + 12 * (j - 1) + 4, 5) = asheet.Cells(2 + 12 * (j - 1) + 4, 5) + co

				asheet.Cells(2 + 12 * (j - 1) + 5, 5) = asheet.Cells(2 + 12 * (j - 1) + 5, 5) + du

				asheet.Cells(2 + 12 * (j - 1) + 9, 5) = asheet.Cells(2 + 12 * (j - 1) + 9, 5) + 1

				'QII,TII,NII

				ElseIf st = "Internal Involuntary Interruptions" Then

				asheet.Cells(2 + 12 * (j - 1) + 1, 5) = asheet.Cells(2 + 12 * (j - 1) + 1, 5) + co

				asheet.Cells(2 + 12 * (j - 1) + 6, 5) = asheet.Cells(2 + 12 * (j - 1) + 6, 5) + du

				asheet.Cells(2 + 12 * (j - 1) + 10, 5) = asheet.Cells(2 + 12 * (j - 1) + 10, 5) + 1

				'QVNB,TVNB,NVNB

				ElseIf st = "Voluntary + Not Budget Interruptions" Then

				asheet.Cells(2 + 12 * (j - 1) + 2, 5) = asheet.Cells(2 + 12 * (j - 1) + 2, 5) + co

				asheet.Cells(2 + 12 * (j - 1) + 7, 5) = asheet.Cells(2 + 12 * (j - 1) + 7, 5) + du

				asheet.Cells(2 + 12 * (j - 1) + 11, 5) = asheet.Cells(2 + 12 * (j - 1) + 11, 5) + 1

				'QN,TVB,NVB

				ElseIf st = "Budget Maintenance Interruptions" Then

				asheet.Cells(2 + 12 * (j - 1) + 3, 5) = asheet.Cells(2 + 12 * (j - 1) + 3, 5) + co

				asheet.Cells(2 + 12 * (j - 1) + 8, 5) = asheet.Cells(2 + 12 * (j - 1) + 8, 5) + du

				asheet.Cells(2 + 12 * (j - 1) + 12, 5) = asheet.Cells(2 + 12 * (j - 1) + 12, 5) + 1

				End If

			End If

		Next j

	End If

Next i

'Blank To Zero

For k = 3 To 86

If asheet.Cells(k, 5) = "" Then

asheet.Cells(k, 5) = 0

End If

Next k

End Sub




