<%
Option Explicit
response.addheader "Content-type", "text/markdown; charset=UTF-8"
sub assert(cond1, cond2, mesg)
    dim text
    if (cond1 = cond2) then
        text = "- [x] " & mesg & vbLf
    else
        text = "- [ ] " & mesg & vbLf
    end if
    response.write text
end sub
%>
<!-- #include file="jsonObject.class.asp" -->
<!-- #include file="underscore.asp" -->

<%
dim U : set U = new Underscore
%>


# Version

<% assert "0.1.2", U.VERSION, "Expected version is 0.1.2." %>


# `.map` Function

Return 10 list items multiplying by 2:

<%
dim c00_index, c00_collection(10)
for c00_index = 0 to ubound(c00_collection)
    dim c00_dictionary : set c00_dictionary = server.createobject("Scripting.Dictionary")
    c00_dictionary.add "index", c00_index
    set c00_collection(c00_index) = c00_dictionary
next

function multiply2(dic, index)
    multiply2 = dic.item("index") * 2
end function

dim c00_results : c00_results = U.map(c00_collection, "multiply2")
%>

<% assert 0, c00_results(0), "Expected value of index 0 is 0." %>
<% assert 2, c00_results(1), "Expected value of index 1 is 2." %>
<% assert 4, c00_results(2), "Expected value of index 2 is 4." %>
<% assert 6, c00_results(3), "Expected value of index 3 is 6." %>
<% assert 8, c00_results(4), "Expected value of index 4 is 8." %>
<% assert 10, c00_results(5), "Expected value of index 5 is 10." %>
<% assert 12, c00_results(6), "Expected value of index 6 is 12." %>
<% assert 14, c00_results(7), "Expected value of index 7 is 14." %>
<% assert 16, c00_results(8), "Expected value of index 8 is 16." %>
<% assert 18, c00_results(9), "Expected value of index 9 is 18." %>


# `.forEach` Sub procedure

Return 10 list items multiplying by 3:

<%
dim c01_index, c01_collection(10)
for c01_index = 0 to ubound(c01_collection)
    dim c01_dictionary : set c01_dictionary = server.createobject("Scripting.Dictionary")
    c01_dictionary.add "index", c01_index
    set c01_collection(c01_index) = c01_dictionary
next

dim c01_results : redim c01_results(ubound(c01_collection))

sub multiply3(dic, index)
    dim num : num = dic.item("index") * 3
    c01_results(index) = num
end sub

U.forEach c01_collection, "multiply3"
%>

<% assert 0, c01_results(0), "Expected value of index 0 is 0." %>
<% assert 3, c01_results(1), "Expected value of index 1 is 3." %>
<% assert 6, c01_results(2), "Expected value of index 2 is 6." %>
<% assert 9, c01_results(3), "Expected value of index 3 is 9." %>
<% assert 12, c01_results(4), "Expected value of index 4 is 12." %>
<% assert 15, c01_results(5), "Expected value of index 5 is 15." %>
<% assert 18, c01_results(6), "Expected value of index 6 is 18." %>
<% assert 21, c01_results(7), "Expected value of index 7 is 21." %>
<% assert 24, c01_results(8), "Expected value of index 8 is 24." %>
<% assert 27, c01_results(9), "Expected value of index 9 is 27." %>


# `.filter` Function

Return 5 of 10 list items filtered by even numbers:

<%
dim c02_index, c02_collection(10)
for c02_index = 0 to ubound(c02_collection)
    dim c02_dictionary : set c02_dictionary = server.createobject("Scripting.Dictionary")
    c02_dictionary.add "index", c02_index
    set c02_collection(c02_index) = c02_dictionary
next

function even(dic, index)
    dim result : result = (dic.item("index") Mod 2 = 0)
    even = result
end function

dim c02_results : c02_results = U.filter(c02_collection, "even")
%>

<% assert 0, c02_results(0).item("index"), "Expected value of index 0 is 0." %>
<% assert 2, c02_results(1).item("index"), "Expected value of index 1 is 2." %>
<% assert 4, c02_results(2).item("index"), "Expected value of index 2 is 4." %>
<% assert 6, c02_results(3).item("index"), "Expected value of index 3 is 6." %>
<% assert 8, c02_results(4).item("index"), "Expected value of index 4 is 8." %>


# `.where` Function

Return 1 of 10 list items filtered by index = 3:

<%
dim c03_index, c03_collection(10)
for c03_index = 0 to ubound(c03_collection)
    dim c03_dictionary : set c03_dictionary = server.createobject("Scripting.Dictionary")
    c03_dictionary.add "index", c03_index
    c03_dictionary.add "test", (c03_index * 2)
    set c03_collection(c03_index) = c03_dictionary
next

dim c03_condition : set c03_condition = server.createobject("Scripting.Dictionary")
c03_condition.add "index", 2
c03_condition.add "test", 4

dim c03_results : c03_results = U.where(c03_collection, c03_condition)
%>

<% assert 2, c03_results(0)("index"), "Expected value of index 0 is 2." %>
<% assert 4, c03_results(0)("test"), "Expected value is 4." %>


Return 1 of 10 list items filtered by index = 3:

<%
dim c04_index
dim c04_jsonarray : set c04_jsonarray = new JSONarray
for c04_index = 0 to 9
    dim c04_jsonobject : set c04_jsonobject = new JSONobject
    c04_jsonobject.add "index", c04_index
    c04_jsonobject.add "test", (c04_index * 2)
    c04_jsonarray.push c04_jsonobject
next

dim c04_condition : set c04_condition = server.createobject("Scripting.Dictionary")
c04_condition.add "index", 2
c04_condition.add "test", 4

dim c04_results : set c04_results = U.where(c04_jsonarray, c04_condition)
%>

<% assert 2, c04_results.itemat(0)("index"), "Expected value of index 0 is 2." %>
<% assert 4, c04_results.itemat(0)("test"), "Expected value is 4." %>
