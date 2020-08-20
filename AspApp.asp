<!DOCTYPE html>
<html>
  <head>
    <title>App</title>
  </head>
  <body>
    <p>
      There
      <%
      If users = 1 Then
        response.write(" is: ")
      Else
        response.write(" are: ")
      End If
      %>
    </p>
    <%
    Dim users
    users=Application("users")
    If users > 0 Then
      response.write("<h1>" & Application("users") & "</h1>")
    Else If users > 1000000 Then
      response.write("<h1 style='color: #FFFF00;'>" & Application("users") & "</h1>")
    Else If users > 1000000000 Then
      response.write("<h1 style='color: #FF0000'>" & Application("users") & "</h1>)
    Else If users > 1000000000000000 Then
      response.write("<h1 style='color: #FF0000'>More than 1000000000000000 users<h1>")
    Else
      response.write("<h1>No</h1>")
    End If
    %>
    <p>active connections in this server.</p>
    
    <script languge="vbscript" runat="server">
      Sub Application_OnStart
        application("vartime")=""
        application("users")=1
      End Sub
    </script>
  </body>
</html>
