<div align="center">

## Net Message


</div>

### Description

To send a Net Message to a user/computer from the web
 
### More Info
 
computer name

I built this to send messages to computer users from a listing of all the computers on the network..can be useful for many things..or just playing with your friends. very simple

message


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stephen King](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephen-king.md)
**Level**          |Beginner
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[System Services/ Functions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/system-services-functions__4-23.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephen-king-net-message__4-6923/archive/master.zip)

### API Declarations

free as a bird


### Source Code

```
<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Send Net Message</title>
</HEAD>
<BODY>
<P align=center>
<% if request("msg") = "" and Request.form("txtmsg") = "" and Request.form("txtmsg2") = "" then
	'detemine if any messages are present; display send link
%>
<A href="message.asp?msg=send" ><STRONG><FONT color=red>Send Message to a computer</FONT></STRONG>  </a>
<%
	else
	if Request.Form("txtmsg") <> "" or Request.form("txtmsg2") <> "" then 'make sure a message has been entered
		if Request("txtcomputer") = "" then 'direct from the listing, not the form
			if request("computer") <> "" then
				Response.Write "<p align=center>Message being sent to "
				Response.Write "<B>" & Request("computer") & "</b>"
				Response.Write ": <font color=blue size=4>"
				Response.Write Request.Form("txtmsg") & "</font></p>"
				set server_shell = Server.CreateObject("wscript.shell")
				server_shell.Run "%comspec% /c net send " & Request("computer") & " " & Request.Form("txtmsg") 'run the command
			else
				Response.Write "<b>No computer entered</b>"
			end if
		else 'the form was used to enter a computer name
			response.Write "<p align=center>Message being sent to "
			Response.Write "<B>" & Request.Form("txtcomputer") & "</b>"
			Response.Write ": <font color=blue size=4>"
			Response.Write Request.Form("txtmsg2") & "</font></p>"
			set server_shell = Server.CreateObject("wscript.shell")
			server_shell.Run "%comspec% /c net send " & Request.Form("txtcomputer") & " " & Request.Form("txtmsg2")
		end if
	elseif request("msg") = "send" then 'draw form for sending message
		if Request("computer") <> "" then 'computer from a listing
			Response.Write "<Form id=frmmsg name=frmmsg method=post action=message.asp>"
			Response.Write "<P align=center>Type in the message you want to send to " & request("computer")
			Response.Write "<input type=text name=txtmsg id=txtmsg length=20>"
			Response.Write "<input type=submit name=submsg id=submsg value=Submit >"
			Response.Write "</form>"
			Response.Write "<br><hr> OR <br>" 'give both options
		end if
		Response.Write "<Form id=frmmsg2 name=frmmsg2 method=post action=message.asp>"
		Response.Write "<P align=center>Type in the message you want to send to "
		'draw a form table
		Response.Write "<table border=0>"
		Response.Write "<tr>"
		Response.Write "<td><b>Computer Name</b></td>"
		Response.Write "<td><input type=text length=20 id=txtcomputer name=txtcomputer></td>"
		Response.Write "</tr><tr>"
		Response.Write "<td><b>Message<b></td>"
		Response.Write "<td><input type=text name=txtmsg2 id=txtmsg2 length=20></td>"
		Response.Write "</tr></table>"
		Response.Write "<input type=submit name=submsg2 id=submsg2 value=Submit >"
		Response.Write "</form>"
	end if
end if
%>
</P>
<P>&nbsp;</P>
</BODY>
</HTML>
```

