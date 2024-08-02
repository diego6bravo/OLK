function setTtlBg(clear)
{
	var custTtlBgL = '<%=Request("custTtlBgL")%>';
	var custTtlBgM = '<%=Request("custTtlBgM")%>';
	var custTtlBg = '';
	if (document.getElementById('tdMyTtl'))
	{
		if (tdMyTtl.length)
		{
			for (var i = 0;i<tdMyTtl.length;i++)
			{
				custTtlBg = (tdMyTtl[i].clientWidth < 450) ? custTtlBgM : custTtlBgL;
				if (!clear)
					tdMyTtl[i].background = '<%=Request("AddPath")%>stretch.aspx?filename=' + custTtlBg + '&w=' + tdMyTtl[i].clientWidth + '&h=' + tdMyTtl[i].clientHeight;
				else
					tdMyTtl[i].background = '';
			}
		}
		else
		{
			custTtlBg = (tdMyTtl.clientWidth > 250 && tdMyTtl.clientWidth < 450) ? custTtlBgM : custTtlBgL;
			if (!clear)
				tdMyTtl.background = '<%=Request("AddPath")%>stretch.aspx?filename=' + custTtlBg + '&w=' + tdMyTtl.clientWidth + '&h=' + tdMyTtl.clientHeight;
			else
				tdMyTtl.background = '';
		}
	}
}
