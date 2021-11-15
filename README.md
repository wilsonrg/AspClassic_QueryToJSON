# ASP Classic 3 + QueryToJSON
##### Converte o resultado do SELECT para o formato JSON

### PASSO 01 - Arquivos/Programas Externos:

**Baixe os arquivos listados abaixo e deixe nas respectivas pastas:**

**_//JQUERY_**

> Link para copiar a versão mais recente: [https://jquery.com/](https://jquery.com/)
> 
> Arquivo: jquery.js ou jquery.min.js
> 
> Pasta: js

**_//JSON 2.0.4_**

>Link para baixar a versão mais recente: <https://gist.github.com/galba/2171058>
>
>Arquivo: json.asp [JSON_2.0.4]
>
>Pasta: json

**_//QueryToJSON_**
>Arquivo em anexo
>
>Arquivo: json_query.asp
>
>Pasta: json
~~~JSON
========== Conteúdo ==========
  'Arquivo do banco de dados conn.asp
  <!--#include file="../bd/conn.asp"-->
  <!--#include file="json.asp"-->
<%
Function QueryToJSON(dbc, sql)
 Dim rs, jsa, col 
 if trim(sql)<>"" then 
    Set rs = dbc.Execute(sql) 
    Set jsa = jsArray() 
    While Not (rs.EOF Or rs.BOF) 
        Set jsa(Null) = jsObject()
        For Each col In rs.Fields
            jsa(Null)(col.Name) = col.Value
        Next
        rs.MoveNext
    Wend 
    Set QueryToJSON = jsa
 else
    QueryToJSON = "" 
 end if 
End Function
%>
========== Conteúdo ==========
~~~

**_//Mysql_**

>Link para baixar a versão mais recente: <https://dev.mysql.com/downloads/mysql/>
>
>Instale o Servidor

**_//IIS_**
>Link para saber como configurar para utilizar o Asp Clássico: <https://www.youtube.com/watch?v=FfSj9VT5nms>

### PASSO 02 - Banco de Dados

**_//TSQL_**
>Criando a Tabela no Banco de Dados
~~~TSQL
CREATE TABLE `login` 
( `id` int(11) NOT NULL AUTO_INCREMENT,
  `login` varchar(15) NOT NULL,
  `senha` varchar(15) NOT NULL,
   PRIMARY KEY (`id`) 
) ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8;

Inserindo informações na tabela login para serem recuperadas depois
INSERT INTO login(login,senha) VALUES('login01','senha01'),('login02','senha02');
~~~

**_//ASP CLASSIC_**
>Crie um arquivo de conexao "conn.asp"
>
>Pasta: bd
~~~ASPCLASSIC
========== Conteúdo ==========
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
========== Conteúdo ==========
~~~

### PASSO 03 - Mãos à Obra

**_//HTML_**
>Crie o arquivo index.html
~~~HTML
<!DOCTYPE html>
<html lang="pt-br">
    <head> 
        <meta charset="UTF-8"> 
        <meta name="viewport" content="width=device-width, initial-scale=1.0"> 
        <meta http-equiv="X-UA-Compatible" content="ie=edge"> 
        <title>HTML com conteúdo Dinâmico</title>
    </head>
    <body> 
        <div id="conteudo"></div> 
        <script language="javascript" type="text/javascript" src="js/jquery.js"></script>
        <script language="javascript" type="text/javascript" src="js/main.js"></script>
    </body>
</html>
~~~

**_//JAVASCRIPT(JQUERY)_**
>Crie o arquivo main.js
>
>Pasta: js
~~~JS
//========== Conteúdo ==========
/*Variáveis globais*/
let v1,v2,v3,v4;
$(function(){ 
    $.getJSON('json/json_query_login.asp',function(data){
        v1='';v2='';v3='';v4='';
        if(data.length > 0){ 
            $.each(data,function(w,itens){ 
                v1 = data[w].id; v1=fv(v1); 
                v2 = data[w].login; v2=fv(v2); 
                v3 = data[w].senha; v3=fv(v3); 
                v4 += '<p>Usuário: '+v2+' - Senha: '+v3+' </p>'; 
            }); 
            $('#conteudo').html(v4); 
        } 
    }); 
});
function fv(v) {
    const valor = v;
    return valor === undefined || valor === null ? '' : valor;
} ========== Conteúdo ==========
~~~

**_//ASP CLASSIC_**
>Cria o arquivo "json_query_login.asp"
>
>Pasta: json
~~~ASP
========== Conteúdo ========== 
 <!--#include file="json_query.asp"--> 
<% 
sql = "select id,login,senha from "&bd_&".login order by login;"
'response.write sql : response.end() 
QueryToJSON(conn,sql).Flush 
%> 
========== Conteúdo ==========
~~~
