var username = document.getElementById("username");
var type_leave = document.getElementById("type_leave");
var datestart = document.getElementById("datestart");
var dateend = document.getElementById("dateend");
var user_selected = document.getElementById("user_selected");
var remaining_leave = document.getElementById("remaining_leave");
var leave_taked = document.getElementById("leave_taked");
var loading = document.getElementById("loading");
var tp = document.getElementById("tp");
var sh1 = document.getElementById("sh1");
var sh2 = document.getElementById("sh2");
var sh3 = document.getElementById("sh3");
var shv = document.getElementById("shv");
var dev = document.getElementById("dev");
var tl = document.getElementById("tl");
var adm = document.getElementById("adm");
var pde = document.getElementById("pde");
var generate_excel = document.getElementById("generate_excel");
var btndownload = document.getElementById("download");
btndownload.disabled = true;
var edit_leave = "n";
var occupation = document.getElementById("occupation");
var info = document.getElementById("info");
var ids = "";
var dj = document.getElementById("demi");
var oj = document.getElementById("one");
var already;
var btnsave = document.getElementById("save_leave");
function getdata(url,id) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        var data = this.responseText.split(";");
        username.innerHTML = data[0] + " " + data[1];
        occupation.innerHTML = data[2];
        ids = data[3];
        tp.innerHTML = data[4];
        remaining_leave.innerHTML = data[5];
        leave_taked.innerHTML = data[6];
        already = JSON.parse(data[7]);
        if(already){
          edit_leave = already._id;
         type_leave.value = already.type.split("(")[0].trim();
         if (already.duration == 0.5){
            pde.style.display = "none";
            datestart.value = already.date_start;
            dj.checked = true;
         }
         else if (already.duration == 1){
          pde.style.display = "none";
          datestart.value = already.date_start;
          oj.checked = true;
         }
         else{
          dj.checked = false;
          oj.checked = false;
          datestart.value = already.date_start;
          dateend.value = already.date_end;
         }
        }
        user_selected.style.display = "block";
      }
    };
    http.send("id="+id);
  }
function define_leave(){
  btnsave.disabled = true;
    if (oj.checked || dj.checked){
      if(oj.checked){
        if (type_leave.value == "" && datestart.value == ""){
          info.innerHTML = "Veuillez remplir tous les informations";
          info.style.display = "block";
     }
     else{
      take_leave("/takeleave",type_leave.value,datestart.value,dateend.value,oj.value,edit_leave);
     }
      }
      else{
        if (type_leave.value == "" && datestart.value == ""){
          info.innerHTML = "Veuillez remplir tous les informations";
          info.style.display = "block";
     }
     else{
      take_leave("/takeleave",type_leave.value,datestart.value,dateend.value,dj.value,edit_leave);
     }
      }
    }
    else{
      if (type_leave.value == "" && datestart.value == "" && dateend.value == ""){
        info.innerHTML = "Veuillez remplir tous les informations";
        info.style.display = "block";
   }
   else{
    take_leave("/takeleave",type_leave.value,datestart.value,dateend.value,"n",edit_leave);
   }
      
    }  
}
function dissapeard(){
  if (dj.checked){
    pde.style.display = "none";
  }
  else{
    pde.style.display = "block";
  }
}
function dissapearo(){
  if (oj.checked){
    pde.style.display = "none";
  }
  else{
    pde.style.display = "block";
  }
}
function take_leave(url,type,startings,endings,val) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        if (this.responseText == "Ok"){
          btnsave.disabled = false;
          if (edit_leave != "n"){
            info.innerHTML = "Congé pour " + username.textContent + " modifier avec succés";
            info.style.display = "block";
            edit_leave = "n";
          }else{
            info.innerHTML = "Congé pour " + username.textContent + " enregistrés";
            info.style.display = "block";
          }
         
        }
        else if (this.responseText == "not authorized"){
          info.innerHTML = username.textContent + " n'est pas autorisée a prendre ce type de congé";
          info.style.display = "block";
        }
        else if (this.responseText == "exceeds"){
          info.innerHTML = username.textContent + " n'a pas assez de solde de congé";
          info.style.display = "block";
        }
        else if (this.responseText == "already"){
          info.innerHTML = username.textContent + " est déja en congé";
          info.style.display = "block";
        }
        else{
          window.location = "/session_end";
        }
      
      }
    };
    http.send("id="+ids+"&type="+type+"&leavestart="+startings+"&leaveend="+endings+"&court="+val+"&edit="+edit_leave);
  }
  function generate(){
    loading.style.display = "block";
    var http = new XMLHttpRequest();
    http.open("POST", "/leave_left", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          btndownload.disabled = false;
          generate_excel.disabled = true;
          loading.style.display = "none";
      }
    }
    http.send();
  }
  function downloads(){
    btndownload.disabled = true;
  }