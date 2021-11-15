<% 
'CABEÇALHO 
response.Charset = "utf-8" 

'DECLARAÇÃO DE VARIÁVEIS 
Dim Servidor,dsnName,dsnUser,dsnPass,database,stringer,bd_,conn 

set conn=Server.CreateObject("ADODB.Connection") 
Servidor = "localhost"  
dsnName = "NomeDSN" 'The name of the DSN 
dsnUser = "USER_BD"  'The username for the DSN 
dsnPass = "SENHA_BD" 'The password for the DSN 
database = "NomeDSN" 'The database to use 'Veja qual driver você tem instalado no seu computador e deixe comentado a linha do driver que você não possui 

stringer = "Provider=MSDASQL;Driver={MySQL ODBC 5.3 ANSI Driver};Server="&Servidor&";Database="&database&";User="&dsnUser&";Password="&dsnPass&";Option=3;" 

stringer = "Driver={MySQL ODBC 3.51 Driver};Server="&Servidor&";Database="&database&";Uid="&dsnUser&";Pwd="&dsnPass&";" 

conn.Open stringer 
'response.write "conexão ok" 
'Alias para query 
bd_ = "NomeDSN" 
%> 
