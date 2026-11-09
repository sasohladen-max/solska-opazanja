import {PublicClientApplication} from "@azure/msal-browser";
const language = document.querySelector("select");
const action = document.querySelector("button");
const copy = {
 sl:{title:"Šolska opažanja",loading:"Priprava prijave …",ready:"Prijavi se s šolskim računom.",success:"Prijava uspešna",failed:"Prijava ni uspela",signIn:"Prijava z Microsoft 365",signOut:"Odjava",busy:"Počakaj na Microsoftovo okno."},
 hu:{title:"Iskolai megfigyelések",loading:"Bejelentkezés előkészítése …",ready:"Jelentkezz be az iskolai fiókoddal.",success:"Sikeres bejelentkezés",failed:"Sikertelen bejelentkezés",signIn:"Bejelentkezés a Microsoft 365-tel",signOut:"Kijelentkezés",busy:"Várj a Microsoft ablakára."}
};
const app = new PublicClientApplication({auth:{clientId:"c4486d58-6a65-4c76-a172-feef9a1aed6f",authority:"https://login.microsoftonline.com/organizations",redirectUri:"https://sasohladen-max.github.io/solska-opazanja/redirect.html"},cache:{cacheLocation:"sessionStorage"}});
let account, problem="", busy=false, initialized=false;
function render(){
 const t=copy[language.value];
 document.documentElement.lang=language.value;
 document.title=document.querySelector("h1").textContent=t.title;
 document.querySelector("[role=status]").textContent=problem?t.failed+" · "+problem:busy?t.busy:!initialized?t.loading:account?t.success+" · "+(account.username || account.name):t.ready;
 action.textContent=account?t.signOut:t.signIn;
 action.disabled=!initialized||busy;
}
language.onchange=render;
action.onclick=async()=>{
 busy=true;problem="";render();
 try{
  if(account){await app.logoutPopup({account,postLogoutRedirectUri:new URL("./",location.href).href});account=null;}
  else {const result=await app.loginPopup({scopes:["openid","profile"],prompt:"select_account"});account=result.account;app.setActiveAccount(account);}
 } catch(error){problem=error.errorCode||error.name;}
 finally{busy=false;render();}
};
render();
try{await app.initialize();account=app.getActiveAccount();initialized=true;}
catch(error){problem=error.errorCode||error.name;}
render();
