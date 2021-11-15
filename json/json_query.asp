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
