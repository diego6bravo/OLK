<script language="javascript">
var GetFormatSep = '<%=Mid(FormatNumber(1000, 2),2,1)%>';
var GetFormatComma = '<%=Mid(FormatNumber(1000, 2),6,1)%>';
function getNumeric(value)
{
	var retVal = value;
	retVal = retVal.replace(GetFormatSep, '');
	retVal = retVal.replace(GetFormatComma, '.');
	return retVal;
}

function getNumericVB(value)
{
	var retVal = value;
	retVal = retVal.replace(GetFormatSep, '');
	return retVal;
}
</script>
