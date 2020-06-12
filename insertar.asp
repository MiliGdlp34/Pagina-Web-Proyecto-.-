<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,nom,email,coments

nom=Request.Form("nombre")
email=request.Form("correo")
coments=Request.Form("coms")

set conn=Server.CreateObject("ADODB.connection")
Conn.open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Sistemas\Music\CARRILLO MILAGROS RODRIGUEZ\Personal.mdb"
Conn.execute "INSERT INTO comentarios(nombre,correo,coments) values('"& nom & "','"& email & "','"& coments & "')"
Conn.close
set conn=nothing
response.redirect("Gracias.html")

%>
</body>
</html>