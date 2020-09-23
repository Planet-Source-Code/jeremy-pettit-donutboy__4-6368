<div align="center">

## Donutboy


</div>

### Description

This code shows the use of a dynamic array for the storage of data. Recordsets are memory hogs, and shouldn't be passed across the server, so why even use one? What I was doing with this code was displaying a list of users on my website. After getting the data, I would have it build a table for the information and create hyperlinks for the appropriate fields. I've seen others do this, but they use recordsets to build them, which defeats its purpose. All the code is done with one trip to the server, so you won't see all those carrot tags being opened and closed throughout my work : )
 
### More Info
 
No parameters are used, no recordsets, no command objects. You do need to create a connection object though. After the connection object is created, you pass it a sql string for the data you want to bring back.

You could create command objects and parameters if you were to use stored procs, but I didn't feel that was necessary for this simple SELECT statement.

I am using ADO for my code, which is pretty much used everywhere, that's why I use it : )

The only thing being brought back is data. It is obtained by executing a SQL string with your connection object also using the GetRows method...

##  vArray = adoCN.Execute(strSQL).GetRows ##

The data is stored in a variable. (which in ASP code, all are variants) A dynamic array will be created. Arrays are zero based, so the first field and first record will be zero.

To get the information you will call it like this...

To get the 2nd record and first field you would use these parameters of your array.

##    vArray(0,1)    ##


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeremy Pettit](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeremy-pettit.md)
**Level**          |Advanced
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeremy-pettit-donutboy__4-6368/archive/master.zip)

### API Declarations

I wrote this code, but I was taught this, and hopefully you learn this too. I don't care if you just steal my code, but a good programmer should know how to do this themselves.


### Source Code

```
<%@ Language=VBScript %>
<%
	DIM vArray
	DIM i
	SUB FillArray
		DIM adoCN
		DIM strSQL
		strSQL = "SELECT UserName, FirstName, LastName, Email, Website FROM Users Order by UserName"
		SET adoCN= server.CreateObject("ADODB.Connection")
		WITH adoCN
			.ConnectionString= "FileDSN=DonutboyWeb"
							.CursorLocation=3
			.Open
			vArray = adoCN.Execute(strSQL).GetRows
		END WITH
		SET adoCN=NOTHING
	END SUB
	response.write "<HTML>"
	response.write "<HEAD>"
	response.write "<META NAME='GENERATOR' Content='Microsoft Visual Studio 6.0'>"
	response.write "</HEAD>"
	response.write "<BODY>"
	response.write "<H1>User Directory</H1>"
	FillArray()
	Response.Write "<TABLE border=1 cellPadding=1 cellSpacing=1 ><TH colspan=2>User Name</TH><TH colspan=2>First Name</TH>"
	Response.Write "<TH colspan=2>Last Name</TH><TH colspan=2>Email</TH><TH colspan=2>Website</TH>"
	for i = 0 to UBOUND(vArray,2)
		Response.Write "<TR>"
		Response.Write "<TD colspan=2>" & vArray(0,i) & "</TD>"
		Response.Write "<TD colspan=2>" & vArray(1,i) & "</TD>"
		Response.Write "<TD colspan=2>" & vArray(2,i) & "</TD>"
		IF instr(1,vArray(3,i),"@") THEN
			Response.Write "<TD colspan=2><A href='mailto:" & vArray(3,i) & "'>" & vArray(3,i) & "</A></TD>"
		ELSE
			Response.Write "<TD colspan=2>" & vArray(3,i) & "</TD>"
		END IF
		IF INSTR(1,vArray(4,i),"http://") THEN
			Response.Write "<TD colspan=2><A href='" & vArray(4,i) & "'>" & vArray(4,i) & "</A></TD>"
		ELSE
			Response.Write "<TD colspan=2>" & vArray(4,i) & "</TD>"
		END IF
		Response.Write "</TR>"
	next
	Response.Write "</TABLE></BODY></HTML>"
%>
```

