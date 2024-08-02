function ValidateForm()
{
   if(document.frmLogin.UserName.value == '') 
   { 
      alert(txtValUser) 
      document.frmLogin.UserName.focus(); 
      return false; 
   } 
 
   if(document.frmLogin.Password.value == '') 
   { 
      alert(txtValPwd) 
      document.frmLogin.Password.focus(); 
      return false; 
   } 
   
   if (document.frmLogin.EnableBranchs.value == 'true' && document.frmLogin.branch.length == 0)
   {
   	  alert(txtValConfBranch);
   	  document.frmLogin.branch.focus();
   	  return false;
   }
   
return true;
 
} 


function focusLogin()
{
	if (!document.frmLogin.UserName.disabled) document.frmLogin.UserName.focus();
}