 <!--#include file="json_query.asp"--> 
<% 
sql = "select id,login,senha from "&bd_&".login order by login;"
'response.write sql : response.end() 
QueryToJSON(conn,sql).Flush 
%> 