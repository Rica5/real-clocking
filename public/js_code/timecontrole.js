function senddata1(){
    var http = new XMLHttpRequest();
    http.open("POST", "/startwork", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
            window.location = "/session_end";
          }
          else{
            
          }
          
      }
    };
    http.send("locaux="+document.getElementById("locaux").value);
}
function senddata2(locauxverif){
    var http = new XMLHttpRequest();
    http.open("POST", "/leftwork", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
            window.location = "/session_end";
          }
          else{
            window.location = "/";
          }
          
      }
    };
    http.send("locaux="+locauxverif);
}
function senddata3(activity){
  var http = new XMLHttpRequest();
  http.open("POST", "/activity", true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
        if (this.responseText == "error"){
           
        }
        else{
        
        }
        
    }
  };
  http.send("activity="+activity);
}