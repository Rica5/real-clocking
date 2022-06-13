var first = document.getElementById("first");
var last = document.getElementById("last");
var m_code = document.getElementById('m_code');
var num_agent = document.getElementById("num_agent");
var btnu = document.getElementById("btnu");
var ids = "";
var textwarn = document.getElementById("textwarn");
var shift = document.getElementById("shift");
btnu.disabled = true;
var del ="";
function getdata(url,id) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        var data = this.responseText.split(",");
        first.value=data[0];last.value=data[1];m_code.value = data[2];num_agent.value = data[3];shift.value=data[4];
        btnu.disabled = false;
        ids = id;
      }
    };
    http.send("id="+id);
  }
  function modify(){
      update_user("/updateuser",ids,m_code.value,num_agent.value,first.value,last.value,shift.value);
  }
  function update_user(url,id,code,num,f,l,shifts) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
         if(this.responseText.includes("already")){
                document.getElementById("notif").setAttribute("style","background-color:red");
                showNotif(this.responseText);
         }
         else if (this.responseText == "error"){
             window.location ="/";
         }
         else{
            document.getElementById("notif").setAttribute("style","background-color:limeagreen");
            showNotif(this.responseText);
         }
      }
    };
    http.send("id="+id+"&code="+code+"&num="+num+"&first="+f+"&last="+l+"&shift="+shifts);
  }
  function showNotif(text) {
    const notif = document.querySelector('.notification');
    notif.innerHTML = text;
    notif.style.display = 'block';
    setTimeout(() => {
        notif.style.display = 'none';
        window.location = "/userlist";
    }, 2000);
  }
  
  function verify(){
      if (m_code.value != "" && num_agent.value != "" && first.value !="" && last.value != "" && shift.value != ""  ){
        btnu.disabled = false;
      } 
      else{
        btnu.disabled = true;
      }
  }

  function delete_user(user){
        textwarn.innerHTML = "Are you sure to delete user <b>"+user+"</b>";
        textwarn.setAttribute("style","color:aliceblue");
        del = user;
  }
  function confirm_del(){
    drop_user("/dropuser",del);
  }
  function drop_user(url,fname) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        if (this.responseText == "error"){
             window.location ="/";
         }
         else{
            showNotif(this.responseText);
         }
      }
    };
    http.send("fname="+fname);
  }
