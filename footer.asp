<table width="90%" border="0">
  <tr>
     
    <td width="22%"> 
      <p style="font-size: 8pt;">Other Nifty Sites:</p></td>
    <td width="78%"><p style="font-size: 8pt;"><b><%

set t = server.CreateObject("adodb.connection")
t.open("dsn=bombness-mysql;uid=bombness;pwd=brianjef")
set q = t.execute("SELECT domain as Domain, sum(hits) as Hits  FROM statsxp.statsxp_referrers  WHERE siteid = 11 and domain <> 'unknown' and domain not in (select data from hshdata where name = 'Domains.Exclude') group by domain  ORDER BY sum(hits) desc limit 5")
while not q.eof
	response.Write("<a href=""http://" & q(0) & """>")
	set r = t.execute("select data from hshdata where name = 'alias:" & lcase(Replace(q(0), "'", "''")) & "' limit 1")
	if r.eof and r.bof then
		response.Write(server.HTMLEncode(q(0)))
	else
		response.Write(server.HTMLEncode(r(0)))
	end if
	r.close
	q.movenext
	response.Write("</a>")
	if not q.eof then response.Write(" | ")
wend
q.close
set q = nothing
set r = nothing
t.close
set t = nothing

	%></b></p></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><p style="font-size: 8pt;">
        <b><a href="/morelinks.asp">More Links</a> | <a href="/morelinks.asp">Your 
        Link Here?</a></b></p></td>
  </tr>
</table>