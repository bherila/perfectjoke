<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
dim conn
set conn = server.CreateObject("adodb.connection")
conn.open("dsn=bombness-mysql;uid=bombness;pwd=brianjef")
set rs = conn.execute("select rating, ratings from jokes where id = '" & Replace(Request("id"), "'", "''") & "' limit 1")
rating = rs(0)
ratings = rs(1)
rs.close
set rs = nothing
dim i
dim r
r = 0.0
for i = 1 to ratings
	r = r + rating
next
ratings = ratings + 1
r = r + cdbl(Request("r"))
r = cdbl(r / ratings)
conn.execute("update jokes set ratings = ratings + 1, rating = '" & r & "' where id = '" & Replace(Request("id"), "'", "''") & "'")
conn.close
set conn = nothing
session("lr") = cint(Request("id"))
response.Redirect("joke.asp?id=" & Request("id"))
%>