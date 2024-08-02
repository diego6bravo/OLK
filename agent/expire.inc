<% 
    pStr = "private, no-cache, must-revalidate" 
    Response.ExpiresAbsolute = #2000-01-01# 
    Response.AddHeader "pragma", "no-cache" 
    Response.AddHeader "cache-control", pStr 
 %>
