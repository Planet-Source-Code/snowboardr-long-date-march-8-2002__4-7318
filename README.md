<div align="center">

## \[ Long Date : March 8, 2002 \]


</div>

### Description

This writes out the full date with the month as a word. If there is a better way to do this let me know. Leave comments if you want.
 
### More Info
 
PWS / IIS


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[snowboardr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/snowboardr.md)
**Level**          |Beginner
**User Rating**    |3.1 (22 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/snowboardr-long-date-march-8-2002__4-7318/archive/master.zip)





### Source Code

```
Function WriteLongDate()
'//By Jason Howard
'//March 8, 2002
Dim ThisMonth
Dim ThisDay
Dim ThisYear
ThisMonth = month(date)
ThisDay = day(date)
ThisYear = year(date)
Select Case ThisMonth
	Case 1
	       ThisMonth = "January"
	Case 2
			ThisMonth = "February"
	Case 3
		 	ThisMonth = "March"
	Case 4
			ThisMonth = "April"
	Case 5
			ThisMonth = "May"
	Case 6
			ThisMonth = "June"
	Case 7
			ThisMonth = "July"
	Case 8
			ThisMonth = "August"
	Case 9
			ThisMonth = "September"
	Case 10
			ThisMonth = "October"
	Case 11
			ThisMonth = "November"
	Case 12
			ThisMonth = "December"
End Select
Response.write(ThisMonth) & " " & ThisDay & ", " & ThisYear
set ThisMonth = Nothing
set ThisDay = Nothing
set ThisYear = Nothing
End Function
'To output this just do this anywhere on your page:
<% WriteLongDate %>
```

