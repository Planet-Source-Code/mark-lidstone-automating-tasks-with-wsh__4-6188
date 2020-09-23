<div align="center">

## Automating Tasks with WSH


</div>

### Description

The purpose of this article is to show how to run tasks at regular intervals without having to access an ASP script with your web browser manually.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Lidstone](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-lidstone.md)
**Level**          |Intermediate
**User Rating**    |4.8 (53 globes from 11 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Server Side](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/server-side__4-31.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-lidstone-automating-tasks-with-wsh__4-6188/archive/master.zip)





### Source Code

<p><b>Automating Tasks</b></p>
<p>A common question with ASP is "How do I run a certain page at regular intervals?". Normally people want to do this because they have a page that performs all their database maintenance or does something like send of reminder emails. Windows NT has a method of scheduling tasks to run at a specific time, but it only allows command-line tasks to be run. To use this all you have to do is use the "AT" command at the console and leave the "Scheduler" service running.</p>
<p>It is possible to start, for instance, Internet Explorer from the command line and tell it to request your page (e.g. "c:\program files\internet explorer\iexplore.exe http://localhost/mypage.asp"). Netscape also gives this ability, but using either will mean that every time the scheduled task runs you will be opening a new browser window and will need to get to the server and actually close the browser window. Running a browser also has a pretty high overhead in terms of disk access, memory space/bandwidth and processor time, and if you went on holiday for a month and the task was run every day you'd come back to a server with about 30 open browser windows. Not nice!</p>
<p>Another drawback is that scripts running through ASP.DLL can timeout. If you are doing a lot of work and you know that the task is going to take longer than your default timeout value, it can make things a little more complicated getting them to run. Although it's possible to make the timeout longer, that means that malformed scripts on other parts of your site can take up more processing time, and changing the timeout for individual pages only allows you to reduce the timeout delay, not lengthen it.</p>
<p>Luckily it is possible to run scripts from the command-line directly, without requesting ASP scripts through the web server thanks to the Windows Scripting Host (WSH). For a rather dry overview of what WSH is, you can see the article at <a href="http://www.microsoft.com/MANAGEMENT/ScrptHost.htm">http://www.microsoft.com/MANAGEMENT/ScrptHost.htm</a></p>
<p>The main advantages to using WSH instead of an ASP script are:
<ol>
	<li>Less memory/CPU intensive than opening a browser.</li>
	<li>Timeouts are optional and can be set on a "per script" basis.</li>
	<li>No windows to close after every execution.</li>
	<li>Simpler code production.</li>
</ol>
<p>Writing WSH scripts is not difficult at all. Normally you can convert your ASP scripts to WSH scripts in a matter of seconds, and to show you what I mean, I'll convert an example ASP script to a WSH script. The source below is for a page that removes all entries in the "tblNewsItems" that are over a week out of date and displays a list of the articles that have been deleted.</p>
<table border=0 width="90%" bgcolor="#EFEFEF" align=center><tr><td width="90%"><pre>
&lt;html&gt;
&lt;head&gt;&lt;title&gt;Database Maintenance Page&lt;/title&gt;
&lt;body background="#FFFFFF"&gt;
&lt;h1&gt;Database Maintenance&lt;/h1&gt;
&lt;!-- #include virtual="/includes/adovbs.inc" --&gt;&lt;%
<br>
	<font color="#007f00">' Define variables</font>
	<font color="#00007f">Dim objConn, objRS
	Dim dtmCutoffDate
	Dim strCutoffDate</font>
<br>
	<font color="#007f00">' Make sure the date format cannot be confused (I'm paranoid about this because I'm British)</font>
	<font color="#00007f">dtmCutoffDate = DateAdd("<font color="#7f0000">d</font>",-7,Date)
	strCutoffDate = Day(dtmCutoffDate) & " " & MonthName(Month(dtmCutoffDate)) & _
		" " & Year(dtmCutoffDate)</font>
<br>
	<font color="#007f00">' Create and setup connection object</font>
	<font color="#00007f">Set objConn = Server.CreateObject("<font color="#7f0000">ADODB.Connection</font>")
	objConn.Open "<font color="#7f0000">MyDSN</font>"</font>
<br>
	<font color="#007f00">' Retrieve records that are to be deleted</font>
	<font color="#00007f">Set objRS = Server.CreateObject("<font color="#7f0000">ADODB.RecordSet</font>")
	objRS.Open "<font color="#7f0000">SELECT * FROM tblNewsItems WHERE dtmExpireDate &lt; #</font>" & strCutoffDate & _
		"<font color="#7f0000">#;</font>", objConn, adOpenKeyset, adLockOptimistic, adCmdText</font>
<br>
	<font color="#007f00">' If there are some articles returned, print their details then remove them from
	' the database</font>
	<font color="#00007f">If NOT objRS.EOF Then
		Response.Write "<font color="#7f0000">The following articles were out of date and have " & _
			been deleted :</font>" & vbCrLf
		Response.Write "<font color="#7f0000">&lt;table border=0 cellpadding=1 cellspacing=1&gt;</font>" & vbCrLf
		Response.Write vbTab & "<font color="#7f0000">&lt;tr&gt;&lt;th&gt;Article Title&lt;/th&gt;&lt;th&gt;Author&lt;/th&gt;" & _
			&lt;th&gt;Start Date&lt;/th&gt;&lt;th&gt;End Date&lt;/th&gt;&lt;/tr&gt;</font>" & vbCrLf
		While NOT objRS.EOF
			Response.Write vbTab & "<font color="#7f0000">&lt;tr&gt;&lt;td&gt;</font>" & objRS("<font color="#7f0000">strTitle</font>") & "<font color="#7f0000">&lt;/td&gt;</font>" & _
				"<font color="#7f0000">&lt;td&gt;</font>" & objRS("<font color="#7f0000">strAuthor</font>") & "<font color="#7f0000">&lt;/td&gt;&lt;td&gt;</font>" & _
				objRS("<font color="#7f0000">dtmStartDate</font>") & "<font color="#7f0000">&lt;/td&gt;&lt;td&gt;</font>" & objRS("<font color="#7f0000">dtmExpireDate</font>") & _
				"<font color="#7f0000">&lt;/td&gt;&lt;/tr&gt;</font>" & vbCrLf
			objRS.MoveNext
		WEnd
		Response.Write "<font color="#7f0000">&lt;/table&gt;</font>" & vbCrLf
		objConn.Execute("<font color="#7f0000">DELETE FROM tblNewsItems WHERE dtmExpireDate &lt; #</font>" & _
			strCutoffDate & "<font color="#7f0000">#;</font>")
	Else</font>
<br>
	<font color="#007f00">' If no out of date articles were found, explain and carry on</font>
		<font color="#00007f">Response.Write "<font color="#7f0000">No out of date articles were found</font>" & vbCrLf
	End If</font>
<br>
	<font color="#007f00">' Tidy up objects</font>
	<font color="#00007f">objRS.Close
	Set objRS = Nothing
	objConn.Close
	Set objConn = Nothing</font>
<br>
%&gt;&lt;/body&gt;
&lt;/html&gt;</pre></td></tr></table>
<p>The main difference is that you don't need to put anything inside script delimiters like "&lt;% .... %&gt;" or "&lt;script runat=Server&gt; .... &lt;/script&gt;" because the entire file is treated as script. Also, there are no "Request" or "Response" objects because there will be no input or output from IIS. Changing our script to take account of this (and deleting everything outside the delimiters) gives us this :</p>
<table border=0 width="90%" bgcolor="#EFEFEF" align=center><tr><td width="90%"><pre>
<font color="#00007f"><font color="#007f00">' Define variables</font>
Dim objConn, objRS
Dim dtmCutoffDate
Dim strCutoffDate
<br>
<font color="#007f00">' Make sure the date format cannot be confused (I'm paranoid about this because I'm British)</font>
dtmCutoffDate = DateAdd("<font color="#7f0000">d</font>",-7,Date)
strCutoffDate = Day(dtmCutoffDate) & " " & MonthName(Month(dtmCutoffDate)) & _
	" " & Year(dtmCutoffDate)
<br>
<font color="#007f00">' Create and setup connection object</font>
Set objConn = Server.CreateObject("<font color="#7f0000">ADODB.Connection</font>")
objConn.Open "<font color="#7f0000">MyDSN</font>"
<br>
<font color="#007f00">' Retrieve records that are to be deleted</font>
Set objRS = Server.CreateObject("<font color="#7f0000">ADODB.RecordSet</font>")
objRS.Open "<font color="#7f0000">SELECT * FROM tblNewsItems WHERE dtmExpireDate &lt; #</font>" & _
	strCutoffDate & "<font color="#7f0000">#;</font>", objConn, adOpenKeyset, adLockOptimistic, adCmdText
<br>
<font color="#007f00">' If there are some articles returned, print their details then remove them from the database</font>
If NOT objRS.EOF Then
	While NOT objRS.EOF
		objRS.MoveNext
	WEnd
	objConn.Execute("<font color="#7f0000">DELETE FROM tblNewsItems WHERE dtmExpireDate &lt; #</font>" & strCutoffDate & "<font color="#7f0000">#;</font>")
Else
<br>
<font color="#007f00">' If no out of date articles were found, explain and carry on</font>
End If
<br>
<font color="#007f00">' Tidy up objects</font>
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing</pre></td></tr></table>
<p>We now have the problem that adovbs.inc is not included, so as a work-around you can open up the adovbs.inc file (or whatever include files you are working with) and copy the relevant lines into your code (luckily adovbs.inc only includes constant definitions and I only need two of them. When using include files that contain large amounts of code this can make the script difficult to navigate) :</p>
<table border=0 width="90%" bgcolor="#EFEFEF" align=center><tr><td width="90%"><pre>
<font color="#00007f"><font color="#007f00">' Define variables</font>
Dim objConn, objRS
Dim dtmCutoffDate
Dim strCutoffDate
<br>
<font color="#007f00">' Define constants from ADOVBS.INC</font>
Const adOpenKeyset = 1
Const adLockOptimistic = 3
Const adCmdText = &H0001
<br>
<font color="#007f00">' Make sure the date format cannot be confused (I'm paranoid about this because I'm British)</font>
dtmCutoffDate = DateAdd("<font color="#7f0000">d</font>",-7,Date)
strCutoffDate = Day(dtmCutoffDate) & " " & MonthName(Month(dtmCutoffDate)) & " " & _
	Year(dtmCutoffDate)
<br>
<font color="#007f00">' Create and setup connection object</font>
Set objConn = Server.CreateObject("<font color="#7f0000">ADODB.Connection</font>")
objConn.Open "<font color="#7f0000">MyDSN</font>"
<br>
<font color="#007f00">' Retrieve records that are to be deleted</font>
Set objRS = Server.CreateObject("<font color="#7f0000">ADODB.RecordSet</font>")
objRS.Open "<font color="#7f0000">SELECT * FROM tblNewsItems WHERE dtmExpireDate &lt; #</font>" & strCutoffDate & _
	"<font color="#7f0000">#;</font>", objConn, adOpenKeyset, adLockOptimistic, adCmdText
<br>
<font color="#007f00">' If there are some articles returned, print their details then remove them from the database</font>
If NOT objRS.EOF Then
	While NOT objRS.EOF
		objRS.MoveNext
	WEnd
	objConn.Execute("<font color="#7f0000">DELETE FROM tblNewsItems WHERE dtmExpireDate &lt; #</font>" & strCutoffDate & "<font color="#7f0000">#;</font>")
Else
<br>
<font color="#007f00">' If no out of date articles were found, explain and carry on</font>
End If
<br>
<font color="#007f00">' Tidy up objects</font>
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing</font></pre></td></tr></table>
<p>In this example we can now see that the object "objRS" is redundant, it's whole point being for display of the data. Taking out all references to that object gives :</p>
<table border=0 width="90%" bgcolor="#EFEFEF" align=center><tr><td width="90%"><pre>
<font color="#00007f">' Define variables
Dim objConn
Dim dtmCutoffDate
Dim strCutoffDate
<br>
<font color="#007f00">' Make sure the date format cannot be confused (I'm paranoid about this because I'm British)</font>
dtmCutoffDate = DateAdd("<font color="#7f0000">d</font>",-7,Date)
strCutoffDate = Day(dtmCutoffDate) & " " & MonthName(Month(dtmCutoffDate)) & " " & _
	Year(dtmCutoffDate)
<br>
<font color="#007f00">' Create and setup connection object</font>
Set objConn = Server.CreateObject("<font color="#7f0000">ADODB.Connection</font>")
objConn.Open "<font color="#7f0000">MyDSN</font>"
<br>
<font color="#007f00">' Run the SQL query</font>
objConn.Execute("<font color="#7f0000">DELETE FROM tblNewsItems WHERE dtmExpireDate &lt; #</font>" & strCutoffDate & "<font color="#7f0000">#;</font>")
<br>
<font color="#007f00">' Tidy up objects</font>
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing</font></pre></td></tr></table>
<p>This can be tidied up even further like so :</p>
<table border=0 width="90%" bgcolor="#EFEFEF" align=center><tr><td width="90%"><pre>
<font color="#00007f">' Define variables
Dim objConn
Dim dtmCutOffDate
<br>
<font color="#007f00">' Make sure the date format cannot be confused (I'm paranoid about this because I'm British)</font>
dtmCutoffDate = DateAdd("<font color="#7f0000">d</font>",-7,Date)
<br>
<font color="#007f00">' Create and setup connection object</font>
Set objConn = Server.CreateObject("<font color="#7f0000">ADODB.Connection</font>")
objConn.Open "<font color="#7f0000">MyDSN</font>"
<br>
<font color="#007f00">' Run the SQL query</font>
objConn.Execute("<font color="#7f0000">DELETE FROM tblNewsItems WHERE dtmExpireDate &lt; #</font>" & Day(dtmCutoffDate) & _
	" " & MonthName(Month(dtmCutoffDate)) & " " & Year(dtmCutoffDate) & "<font color="#7f0000">#;</font>")
<br>
<font color="#007f00">' Tidy up objects</font>
objConn.Close
Set objConn = Nothing</font></pre></td></tr></table>
<p>The script has now been stripped down to the bare basics with nothing except the real bones functionality of the original. This is much cleaner to look at and will be more efficient.</p>
<p>Out of interest, it is also possible to remove references to the connection object instead of references to the recordset object like so :</p>
<table border=0 width="90%" bgcolor="#EFEFEF" align=center><tr><td width="90%"><pre>
<font color="#00007f"><font color="#007f00">' Define variables</font>
Dim objRS
<br>
<font color="#007f00">' Define constants</font>
Const adOpenKeyset = 1
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const adAffectAll = 3
<br>
<font color="#007f00">' Create and setup recordset object</font>
objRS.Open "<font color="#7f0000">SELECT * FROM tblNewsItems WHERE dtmExpireDate &lt; #</font>" & strCutoffDate & _
	"<font color="#7f0000">#;</font>", "<font color="#7f0000">MyDSN</font>", adOpenKeyset, adLockOptimistic, adCmdText
<br>
<font color="#007f00">' Remove out of date articles from the database</font>
objRS.Delete adAffectAll
objRS.Update
<br>
<font color="#007f00">' Tidy up objects</font>
objRS.Close
Set objRS = Nothing</font></pre></td></tr></table>
<p>but this method is slightly less efficient because it returns the matching records before deleting them from the database whereas using the connection object deletes the entries directly from the database without loading them into memory first.</p>
<p>Alternatively, you may find it easier to just rewrite your code from scratch. This way you shouldn't end up accidentally including some code that was meant for formatting output which isn't needed any more.</p>
<p>Now that you have a script like this you need to save it with the ".vbs" extension. If you now look at the script in an Explorer window it shouw have an icon like a small scroll of blue paper. Double-clicking it actually runs the script and performs the same funcion as the ASP page on your website, but without the need to access it with a browser.</p>
<p>If you go to the command-line and try typing in the name of the file, you will get the standard "I don't know what to do with this file" message that you get from the console, so what you need to do is tell it that you want to run the WSH and pass it the script.</p>
<p>There are two ways to call the WSH engine from the command-line, which are "CSCRIPT" which calls the command-line version of the WSH, and "WSCRIPT" which calls the windows version. As the command-line version seems to have a lower overhead I'll stick to that one. Let's assume that your script is saved as "dbmaintain.vbs" in the "c:\scripts\" directory. The console command to run that script would be "cscript c:\scripts\dbmaintain.vbs". You can pass this command directly to the AT scheduller, or you can place it in a batch file and pass the batch file to AT. Voila! You now have a working maintenance script.</p>
<p>Here's a quick tip. If you have a server that has several tasks that need to be run at regular intervals you might find it easier to create a set of batch files representing different time-plans or repetition frequencies. e.g. you could have a batch file called "hour.bat" which is run every hour, one called "day.bat" which is run every day etc... This means you don't have to re-type the long AT commands every time you want to add or change a task, and it makes looking up what tasks are run at what frequency much easier.</p>
<p>Of course, database maintenance is not the only thing that this can be useful for. Several people have asked if it is possible to do something like email a client x number of days before an advertisement they have placed expires. This again is simple and just needs the SQL statement to be changed to something like:</p>
<table border=0 width="90%" bgcolor="#EFEFEF" align=center><tr><td width="90%"><pre>strSQL = "SELECT strPlacerName, strPlacerEmail, dtmExpireDate FROM " & _
	"tblAdverts WHERE dtmExpireDate &gt;= #" & DateAdd("d",-5,Date) & "#;"</pre></td></tr></table>
<p>Which would return a recordset populated with all the adverts that are due to expire in the next 5 days. Emailing all of these people then would be a simple matter using CDONTS or some other mailer component. You could then improve the system to query the database for anyone whose advert expires in 5 days and give them a reminder, then look for ads that expire in 2 days and give them a more urgent message etc....</p>

