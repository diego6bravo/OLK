function changeDB()
{
	$.post('changeDB.asp', { dbID: document.Form1.dbID.value }, function(data)
	{
		var arrData = data.split('{S}');
		document.Form1.EnableBranchs.value = arrData[0];
		
		var Branchs = document.Form1.branch
		for (var i = Branchs.length-1;i>=0;i--)
		{
			Branchs.remove(i)
		}
		
		var txtUserName = document.getElementById('UserName');
		var txtPassword = document.getElementById('Password');
		var btnEnter = document.getElementById('btnEnter');
		var trUpdate = document.getElementById('trUpdate');
		var trBranch = document.getElementById('trBranch');
		var chkSave = document.getElementById('Save');
		var EnableBranchs = document.getElementById('EnableBranchs');
		
		if (arrData[0] == 'Y')
		{
			EnableBranchs.value = arrData[1];
			if (arrData[1] == 'Y')
			{
				var arrBranch = arrData[2].split('{O}');
				for (var i = 0;i<arrBranch.length;i++)
				{
					var arrBranchData = arrBranch[i].split('{C}');
					Branchs.options[i] = new Option(arrBranchData[1], arrBranchData[0]);
				}
				trBranch.style.display = '';
			}
			else
			{
				trBranch.style.display = 'none';
			}
			
			txtUserName.disabled = false;
			txtPassword.disabled = false;
			btnEnter.disabled = false;
			if (chkSave) chkSave.disabled = false;
			trUpdate.style.display = 'none';
		}
		else
		{
			txtUserName.disabled = true;
			txtPassword.disabled = true;
			btnEnter.disabled = true;
			if (chkSave) chkSave.disabled = true;
			trUpdate.style.display = '';
			trBranch.style.display = 'none';
		}
	});
}

function ValidateForm()
{
   if(document.Form1.UserName.value == '') 
   { 
      alert(txtValUser) 
      document.Form1.UserName.focus(); 
      return false; 
   } 
 
   if(document.Form1.Password.value == '') 
   { 
      alert(txtValPwd) 
      document.Form1.Password.focus(); 
      return false; 
   } 
   
   if (document.Form1.EnableBranchs.value == 'Y' && document.Form1.branch.length == 0)
   {
   	  alert(txtValConfBranch);
   	  document.Form1.branch.focus();
   	  return false;
   }
   
return true;
 
} 
