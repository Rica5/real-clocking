var username = document.getElementById("username");
var type_leave = document.getElementById("type_leave");
var datestart = document.getElementById("datestart");
var dateend = document.getElementById("dateend");
var user_selected = document.getElementById("user_selected");
var remaining_leave = document.getElementById("remaining_leave");
var leave_taked = document.getElementById("leave_taked");
var tp = document.getElementById("tp");
var sh1 = document.getElementById("sh1");
var sh2 = document.getElementById("sh2");
var sh3 = document.getElementById("sh3");
var dev = document.getElementById("dev");
var tl = document.getElementById("tl");
var adm = document.getElementById("adm");
var occupation = document.getElementById("occupation");
var info = document.getElementById("info");
var ids = "";
function getdata(url,id) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        var data = this.responseText.split(",");
        username.innerHTML = data[0] + " " + data[1];
        occupation.innerHTML = data[2];
        ids = data[3];
        tp.innerHTML = data[4];
        remaining_leave.innerHTML = data[5];
        leave_taked.innerHTML = data[6];
        user_selected.style.display = "block";
      }
    };
    http.send("id="+id);
  }
function define_leave(){
       if (type_leave.value == "" && datestart.value == "" && dateend.value == ""){
            info.innerHTML = "Please fill all informations";
            info.style.display = "block";
       }
       else{
        take_leave("/takeleave",type_leave.value,datestart.value,dateend.value);
       }
}
function take_leave(url,type,startings,endings) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        if (this.responseText == "Ok"){
          info.innerHTML = "Leave from work for " + username.textContent + " registered";
          info.style.display = "block";
        }
        else if (this.responseText == "not authorized"){
          info.innerHTML = username.textContent + " not authorized to take this type of live from work";
          info.style.display = "block";
        }
        else if (this.responseText == "exceeds"){
          info.innerHTML = username.textContent + " have not enough remaining leave";
          info.style.display = "block";
        }
        else if (this.responseText == "already"){
          info.innerHTML = username.textContent + " is already on leave from work";
          info.style.display = "block";
        }
        else{
          window.location = "/session_end";
        }
      
      }
    };
    http.send("id="+ids+"&type="+type+"&leavestart="+startings+"&leaveend="+endings);
  }