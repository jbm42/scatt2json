on error resume next

Dim fso
Set fso = CreateObject ("Scripting.FileSystemObject")

Dim stdout
Set stdout = fso.GetStandardStream (1)

Dim doc
Set doc = CreateObject("ScattDoc.ScattDocument")
if doc Is Nothing then
	stdout.WriteLine "Error: Scatt Professional was not installed properly."
	return
end if

Set objArgs = WScript.Arguments
If objArgs.Count > 0 Then
        Dim i
	For i = 0 to objArgs.Count - 1
		ProcessDocument doc, objArgs(i)
	Next
Else
	set folder = fso.GetFolder(".")

	for each file in folder.files
		if right(file,6)=".scatt" then
	        ProcessDocument doc, file
		end if
	next
End If

Sub ProcessDocument(doc,filename)
	doc.FileName = filename

	doc.Load
	if Not doc.Valid then
		stdout.WriteLine "Error: File not found or invalid file format."
		return
	end if

	Dim AllShots
	Set AllShots = doc.Aimings.Match.Shots
	if AllShots.Count = 0 then
		stdout.WriteLine "Error: No match shots"
		return
	end if

	Dim outputFilename
	Dim d
	d = AllShots(1).ShotTime
	d = mid(d,7,4) & mid(d,4,2) & left(d,2) & "_" & mid(d,12,2) & mid(d,15,2) & mid(d,18,2)

	outputFilename = doc.ShooterName & " " & doc.Event.ShortName & " " & d
	outputFilename = replace(outputFilename," ","_")
	outputFilename = replace(outputFilename,":","") & ".json"
	stdout.WriteLine filename & " => " & outputFilename

	Dim outputFile
	Set outputFile = fso.CreateTextFile(outputFilename , True)
	if Err.Number <> 0 then
	 	stdout.WriteLine "Error: Can't create output file - " & Err.Description
		return
	end if

	outputFile.WriteLine "{"
	outputFile.WriteLine "  ""event_name"": """ & doc.Event.Name & ""","
	outputFile.WriteLine "  ""event_short_name"": """ & doc.Event.ShortName & ""","
	outputFile.WriteLine "  ""caliber"": """ & doc.Event.Caliber & ""","
	outputFile.WriteLine "  ""shooter"": """ & doc.ShooterName & ""","
	outputFile.WriteLine "  ""date"": """ & AllShots(1).ShotTime & ""","
	outputFile.WriteLine "  ""comments"": """ & doc("_comments") & ""","

	Dim comma

	Dim Targ
	Set Targ = doc.Event.Target(1)

	outputFile.WriteLine "  ""rings"": ["
	for i=1 to 10
		if i=10 then comma="" else comma=","

		outputFile.WriteLine "    " & Targ.Ring(i) & comma
	next
	outputFile.WriteLine "  ],"

	outputFile.WriteLine "  ""shots"": ["

	Dim i
	For i = 1 To AllShots.Count
		Dim CurrentShot
		Set CurrentShot = AllShots(i)

		outputFile.WriteLine "    {"
		outputFile.WriteLine "      ""number"": " & CStr(i) & ","
		outputFile.WriteLine "      ""f_coefficient"": """ & CurrentShot.Attr.FCoefficient & ""","
		outputFile.WriteLine "      ""enter_time"": """ & CurrentShot.Attr.EnterTime & ""","
		outputFile.WriteLine "      ""shot_time"": """ & CurrentShot.ShotTime & ""","
		outputFile.WriteLine "      ""result"": " & CurrentShot.Result & ","
		outputFile.WriteLine "      ""breach_x"": " & FormatNumber(CurrentShot.BreachX, 2) & ","
		outputFile.WriteLine "      ""breach_y"": " & FormatNumber(CurrentShot.BreachY, 2) & ","
		outputFile.WriteLine "      ""trace"": ["

		dim TestRange
		Set TestRange = CurrentShot.Range(-10, 10)

		Dim j
		For j = TestRange.First to TestRange.Last
	    if j=TestRange.Last then comma="" else comma=","

			outputFile.WriteLine "        { ""t"": " & FormatNumber(TestRange.Index2Sec(j), 3) _
				& ", ""x"": " & FormatNumber(TestRange.X(j), 2) _
	                        & ", ""y"": " & FormatNumber(TestRange.Y(j), 2) & " }" & comma
		Next

		if i=AllShots.Count then comma="" else comma=","

		outputFile.WriteLine "      ]"
		outputFile.WriteLine "    }" & comma
	Next
	outputFile.WriteLine "  ]"
	outputFile.WriteLine "}"
	outputFile.Close
End Sub
