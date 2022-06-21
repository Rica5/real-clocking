var w = document.getElementById("w");
var a = document.getElementById("a");
var l = document.getElementById("l");
var s = document.getElementById("s");
var info = document.getElementById("info");
var clock = document.getElementById("time");
var chloc = document.getElementById("locaux");
a.disabled = true;
l.disabled = true;
var day = ["Dimanche","Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi"];
function workings(){
    a.disabled = false;
    l.disabled = false;
    w.disabled = true;
    s.setAttribute("style","background:#57b846;font-size:12px;");
    s.innerHTML = "TRAVAILLER";
    chloc.style.display = "none";
    info.style.display = "block";
    info.innerHTML = "Lieu de travail : "+ chloc.value;
}

function aways(){
    w.disabled = false;
    l.disabled = false;
    a.disabled = true;
    s.setAttribute("style","background:#FFBA00;font-size:12px;");

    s.innerHTML = "ABSENT(E)";
    chloc.style.display = "none";
    info.style.display = "block";
    info.innerHTML = "Lieu de travail "+ chloc.value;
}

function lefts(){
    w.disabled = false;
    a.disabled = false;
    l.disabled = true;
    chloc.style.display = "block";
    info.style.display = "none";
    chloc.value = "Not defined";
    s.setAttribute("style","background:#E53F31;font-size:12px;");
    s.innerHTML = "PARTI";
    w.style.display = "none";
		a.style.display = "none";
		l.style.display = "none";
}
function clockRunner () {
    var datetime = day[new Date().getDay()]+ "  a  "+ new Date().toLocaleTimeString();
    clock.innerHTML = datetime;
    setTimeout(clockRunner, 1000);
  }
  
  window.onload = clockRunner();
 
