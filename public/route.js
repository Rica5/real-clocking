const express = require("express");
const routeExp = express.Router();
const mongoose = require("mongoose");
const UserSchema = require("../models/User");
const StatusSchema = require("../models/status");
const LateSchema = require("../models/late");
const AbsentSchema = require("../models/absent");
const LeaveSchema = require("../models/leave");
const nodemailer = require("nodemailer");
const crypto = require('crypto');
const moment = require("moment");
const ExcelFile = require("sheetjs-style");
const fs = require('fs');
var date_data = [];
var data = [];
var all_datas = [];
var num_file = 1;
var hours = 0;
var minutes = 0;
var notification = [];

//Mailing
var transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: "ricardoramandimbisoa@gmail.com",
    pass: "eisproxichxdutzk",
  },
});

//Page login
routeExp.route("/").get(async function (req, res) {
    session = req.session;
    if (session.occupation_u == "User") {
      res.redirect("/employee");
    } else if (session.occupation_a == "Admin") {
      res.redirect("/home");
    } else {
      await checkleave();
      res.render("Login.html", { erreur: "" });
    }
});

routeExp.route("/latelist").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var latelist = await LateSchema.find({validation:true,reason:{$ne:""}});
      res.render("latelist.html",{latelist:latelist,notif:notification});
    });
  }
  else{
  res.redirect("/");
  }
});
//Login post
routeExp.route("/login").post(async function (req, res) {
  session = req.session;
  await login(req.body.username,req.body.pwd,session,res);
});
routeExp.route("/getip").post(async function (req, res) {
  session = req.session;
  session.ip = req.body.ip;
  res.send("Ok");
});
async function login(username,pwd,session,res){
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      let hash = crypto.createHash('md5').update(pwd).digest("hex");
      var logger = await UserSchema.findOne({
        username: username,
        password: hash,
      });
      if (logger) { 
        //Tete
        if ((logger.shift == "SHIFT 1" || logger.shift == "SHIFT 2") && ((session.ip != "102.16.26.233" && session.ip != "102.16.26.115" && session.ip != "41.63.146.186"))){
          res.render("denied.html");
        }
        else{
        if (logger.change != "n"){
          if (logger.occupation == "User") {
            session.forget = "n";
            if (await StatusSchema.findOne({m_code:logger.m_code,time_end:"",date:{$ne:moment().format("YYYY-MM-DD")}})){
                session.forget ="y";
                await UserSchema.findOneAndUpdate({m_code:logger.m_code},{act_stat:"LEFTING",act_loc:"Not defined"});
            }
            session.occupation_u = logger.occupation;
            session.m_code = logger.m_code;
            session.shift = logger.shift;
          if (difference_year(logger.save_at) && logger.remaining_leave == 0){
              await leave_permission(logger.m_code);
          }
          if (logger.act_stat == "VACATION"){
            session.occupation_u = null;
            session.m_code = null;
            session.shift = null;
            res.render("Login.html", {
              erreur: "Vous ??tes en cong?? prenez votre temps",
            });
          }
          else{
            if (await StatusSchema.findOne({m_code:session.m_code,date:moment().format("YYYY-MM-DD")})){
              session.late = "y";
              await UserSchema.findOneAndUpdate({m_code:session.m_code},{late:"y"});
            }
            else{
              session.late = "n";
              await UserSchema.findOneAndUpdate({m_code:session.m_code},{late:"n"});
            }
            session.name = logger.first_name + " " + logger.last_name;
            session.num_agent = logger.num_agent;
            var late = await LateSchema.findOne({m_code:logger.m_code,date:moment().format("YYYY-MM-DD"),reason:""});
            
            if (late){
              session.time = late.time + " minutes";
              res.redirect("/employee");
          }
          else{
            var already = await LateSchema.findOne({m_code:logger.m_code,date:moment().format("YYYY-MM-DD")});
           
            if (already){
              session.time = "y";
              res.redirect("/employee");
            }
            else{
              if (session.late == "n"){
              var start="";
              var today = moment().day();
            switch(session.shift){
              case "SHIFT 1": start = "06:15";break;
              case "SHIFT 2": start = "12:15";break;
              case "SHIFT 3": start = "18:15";break;
              case "DEV"    : start = "08:00";break;
              case "RH": start = "08:00";break;
              case "MANAGER": start = "08:00";break;
              case "TL": start = "06:15";break;
              case "GERANT" : start = "08:00";break;
              case "ENG" : start = "08:00";break;
              case "IT" : start = "06:15";break;
              default: start = "08:00";break;
            }
            switch(today){
              case 6 : start="08:00";break;
              case 7 : start="08:00";break;
              default: start = "08:00";break;
            }
            start = "20:00";
            var timestart = moment().add(3,'hours').format("HH:mm");
            var time = calcul_retard(start,timestart);
              if ( time > 10){
                var new_late = {
                  m_code:session.m_code,
                  num_agent : session.num_agent,
                  date:moment().format("YYYY-MM-DD"),
                  nom:session.name,
                  time:time,
                  reason:"",
                  validation:false
                }
                await LateSchema(new_late).save();
                session.time = time + " minutes";
               res.redirect("/employee");
            }
            else{
              session.time = "y";
              res.redirect("/employee");
            }
            await UserSchema.findOneAndUpdate({m_code:session.m_code},{late:"y"});
          }
          else{
            session.time = "y";
            res.redirect("/employee");
          }
            }
          }
          }
          } else {
             session.occupation_a = logger.occupation;
             if (notification.length > 16 ){
               for(n=0;n < 8;n++){
                 delete notification[0];
               }
             }
             session.request = {};
            res.redirect("/home");
          }
        }
        else{
          session.mailconfirm = logger.username;
          res.render("change_password.html", {
            first: "y",
          });
        }
      }
       //Pied
      } else {
        res.render("Login.html", {
          erreur: "Email ou mot de passe incorrect",
        });
      }
    });
}
//Validation page
routeExp.route("/validelate").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    session.filtrage = null;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var latelist = await LateSchema.find({ validation: false});
        res.render("latevalidation.html", { latelist: latelist,notif:notification });
      });
  } else {
    res.redirect("/");
  }
});
//Validation
routeExp.route("/validate").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
  var id = req.body.id;
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await LateSchema.findOneAndUpdate(
        { _id: id },
        { validation: true }
      );
      res.send("Ok");
    });
  }
  else{
  res.redirect("/");
  }
});
//Denied
routeExp.route("/denied").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
  var id = req.body.id;
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await LateSchema.findOneAndDelete({ _id: id });
      res.send("Ok");
    });
  }
  else{
    res.send("retour");
  }
});
routeExp.route("/reason").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User"){
    await reason_late(req.body.reason,session,res);
  }
  else{
    res.redirect("/")
  } 
})
routeExp.route("/gethour").post(async function (req, res) {
  session = req.session;
  var ht = req.body.hour;
  if (session.occupation_u == "User"){
    mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
      await UserSchema.findOneAndUpdate({m_code:session.m_code},{user_ht:ht});
      res.send("Ok");
  });
  }
  else{
    res.redirect("/")
  } 
})
//reason late 
async function reason_late(reason,session,res){
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
    session.time = "y";
   await LateSchema.findOneAndUpdate({m_code:session.m_code,reason:""},{reason:reason});
   res.send("Ok");
  })
}
routeExp.route("/employee").get(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User"){
    mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var user = await UserSchema.findOne({m_code:session.m_code});
      if (session.time != "y"){
        res.render("Working.html",{user:user,retard:session.time,forget:session.forget});
      }
      else{
        res.render("Working.html",{user:user,retard:"n",forget:session.forget});
      }
   
    })
  }
  else{
    res.redirect("/")
  } 
})
routeExp.route("/newuser").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin"){
  res.render("newuser.html");
  }
  else{
    res.redirect("/")
  }
})
//Absent
routeExp.route("/absent").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User"){
      await absent(req.body.reason,req.body.stat,session,res);
  }
  else{
    res.send("error");
  }
});
//Absent
async function absent(reason,stat,session,res){
  var timestart = moment().add(3,'hours').format("HH:mm");
  var new_abs = {
    m_code:session.m_code,
    num_agent : session.num_agent,
    date:moment().format("YYYY-MM-DD"),
    nom:session.name,
    time_start:timestart,
    return:"Not come back",
    reason:reason,
    validation:false,
    status:"none"
  }
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await AbsentSchema(new_abs).save();
      res.send(stat);
    });
}
routeExp.route("/absencelist").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin"){
    mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var absent = await AbsentSchema.find({validation:true});
      res.render("abscencelist.html",{absent:absent,notif:notification});
    })
  }
  else{
    res.redirect("/");
  }
})
//Validation page
routeExp.route("/valideabsence").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    session.filtrage = null;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var absent = await AbsentSchema.find({ validation: false});
        res.render("absencevalidation.html", { absent: absent,notif:notification });
      });
  } else {
    res.redirect("/");
  }
});
//Validation
routeExp.route("/validate_a").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
  var id = req.body.id;
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await AbsentSchema.findOneAndUpdate(
        { _id: id },
        { validation: true,status:"accepted"}
      );
      res.send("Ok");
    });
  }
  else{
  res.redirect("/");
  }
});
//Denied
routeExp.route("/denied_a").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
  var id = req.body.id;
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await AbsentSchema.findOneAndUpdate({ _id: id },{status:"denied"});
      res.send("Ok");
    });
  }
  else{
    res.send("retour");
  }
});
routeExp.route("/forget").post(async function (req, res) {
  session = req.session;
  var tf = req.body.timeforget;
  if (session.occupation_u == "User") {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var user = await UserSchema.findOne({m_code:session.m_code});
      var last_time = await StatusSchema.findOne({m_code:session.m_code,time_end:""});
      if (parseInt(user.user_ht) != 0){
        if (hour_diff(last_time.time_start,tf) <= (parseInt(user.user_ht) + 1)){
          await StatusSchema.findOneAndUpdate({m_code:session.m_code,time_end:""},{time_end:tf});
          session.forget = "n";
           res.send("Ok");
        }
        else{
          res.send("No");
        }
      }
      else{
        await StatusSchema.findOneAndUpdate({m_code:session.m_code,time_end:""},{time_end:tf});
        session.forget = "n";
         res.send("Ok");
      }
     
    
    });
  }
  else{
    res.send("retour");
  }
});
//activity
routeExp.route("/activity").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User"){
    await activity(req.body.activity,session,req,res);
  }
  else{
    res.send("error");
  }
  
});
async function activity(activity,session,req,res){
  if (activity != "ABSENT"){
  var counter = 0;
  switch(activity){
    case "BREAK" : counter = 1200000;session.place="petit break";break;
    case "DEJEUNER" : counter = 2400000;session.place = "D??jeuner";break;
    case "PAUSE" : counter = 1800000;session.place = "Pause";break;
  }
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
    await UserSchema.findOneAndUpdate({m_code:session.m_code},{$inc : {count:1},take_break:"n"});
    var counts = await UserSchema.findOne({m_code:session.m_code});
    if (counts.count > 6){
      notification.push(session.name + " quitte son poste trop frequement ");
      const io = req.app.get('io');
      io.sockets.emit('notif', notification );
    }
    setTimeout(async function() {
      counts = await UserSchema.findOneAndUpdate({m_code:session.m_code},{$inc : {count:1},take_break:"n"});
      if (counts.take_break == "n"){
        notification.push(session.name + " prend trop de temp au " + session.place);
        const io = req.app.get('io');
        io.sockets.emit('notif', notification );
      }
     },counter);
  })
  res.send("Ok");
}
else{
  res.send("Ok");
}

} 
routeExp.route("/takebreak").post(async function (req, res) {
  session = req.session;
  await take_break(session);
  res.send("Ok");
})
async function take_break(session){
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
    await UserSchema.findOneAndUpdate({m_code:session.m_code},{take_break:"y"});
  })
}
routeExp.route("/home").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin"){
    mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      res.render("dashboard.html",{notif:notification});
    });
  }
  else{
    res.redirect("/");
  }
})
//Details
routeExp.route("/details").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin"){
    mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      if (session.filtrage == ""){
        res.render("details.html",{timesheets:session.datatowrite,notif:notification});
      }
      else{
        var timesheets =  await StatusSchema.find({time_end:{ $ne: "" }});
        session.datatowrite = timesheets;
        session.datalate = await LateSchema.find({validation:true});
        session.dataabsence = await AbsentSchema.find({validation:true});
        session.dataleave = await LeaveSchema.find({});
        res.render("details.html",{timesheets:timesheets,notif:notification});
      }  
    })
  }
  else{
    res.redirect("/");
  }
  })

routeExp.route("/management").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin"){
    session.filtrage = null;
    mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var alluser = await UserSchema.find({});
      res.render("status.html",{users:alluser,notif:notification});
    })
  }
  else{
    res.redirect("/");
  }
})
routeExp.route("/changepassword").get(async function (req, res) {
  res.render("change_password.html",{first:""});
});
//getuser
routeExp.route("/getuser").post(async function (req, res) {
  var id = req.body.id;
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
    var user = await UserSchema.findOne({_id:id });
    res.send(user.first_name +","+user.last_name+","+user.m_code + ","+user.num_agent +","+user.shift);
  });
});
//getuser
routeExp.route("/getuser_leave").post(async function (req, res) {
  var id = req.body.id;
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
    var user = await UserSchema.findOne({_id:id });
    res.send(user.first_name +","+user.last_name+","+user.shift+","+user._id + ","+time_passed(user.save_at)+","+user.remaining_leave +","+user.leave_taked);
  });
});
//Update User
routeExp.route("/updateuser").post(async function (req, res) {
  var id = req.body.id;
  var m_code = req.body.code;
  var num_agent = req.body.num;
  var amount = req.body.am;
  var first = req.body.first;
  var last = req.body.last;
  var shift = req.body.shift;
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
    var user = await UserSchema.findOne({_id:id });
      await StatusSchema.updateMany({m_code:user.m_code},{m_code:m_code,num_agent:num_agent,nom:first + " "+ last});
      await UserSchema.findOneAndUpdate({_id:id },{m_code:m_code,num_agent:num_agent,amount:amount,first_name:first,last_name:last,shift:shift});
      res.send("User updated successfully");
})
})
//Drop user 
routeExp.route("/dropuser").post(async function (req, res) {
  var names = req.body.fname;
  names = names.split(" ");
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
      await UserSchema.findOneAndDelete({first_name:names[0],last_name:names[1]});
      res.send("User deleted successfully");
});
});
routeExp.route("/userlist").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin"){
    session.filtrage = null;
    mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var alluser = await UserSchema.find({});
      res.render("userlist.html",{users:alluser,notif:notification});
    })
  }
  else{
    res.redirect("/");
  }
})

//Change password
routeExp.route("/checkmail").post(async function (req, res) {
  session = req.session
  var email = req.body.email;
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      if (await UserSchema.findOne({ username: email })) {
        session.mailconfirm = email;
        session.code = randomCode();
        sendEmail(
          session.mailconfirm,
          "Verification code",
          htmlVerification(session.code)
        );
        res.send("done");
      } else {
        res.send("error");
      }
    });
})

// 
routeExp.route("/checkcode").post(async function (req, res) {
  session = req.session;
  if (session.code == req.body.code) {
    res.send("match");
  } else {
    res.send("error");
  }
});
routeExp.route("/changepass").post(async function (req, res) {
  session = req.session
  var newpass = req.body.pass;
  let hash = crypto.createHash('md5').update(newpass).digest("hex");
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      if (await UserSchema.findOne({username: session.mailconfirm,password:hash})){
        res.send("error")
      }
      else{
        await UserSchema.findOneAndUpdate(
          { username: session.mailconfirm },
          { password: hash,change:"y" }
        );
        session.mailconfirm = null;
        session.code = null;
        res.send("Ok");
      }
    });
});
routeExp.route("/startwork").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User"){
    await startwork(req.body.locaux,session,res);
  }
  else {
    res.send("error");
  }
});
//startwork
async function startwork(locaux,session,res){
  var date = moment().format("YYYY-MM-DD");
  var timestart = moment().add(3,'hours').format("HH:mm");
  var new_time = {
    m_code:session.m_code,
    num_agent : session.num_agent,
    date:date,
    time_start:timestart,
    time_end:"",
    nom:session.name,
    locaux:locaux
  }
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
        await StatusSchema(new_time).save();
        if (await StatusSchema.findOne({m_code:session.m_code,date:moment().format("YYYY-MM-DD")})){
            await AbsentSchema.findOneAndUpdate({m_code:session.m_code,return:"Not come back",date:moment().format("YYYY-MM-DD")},{return:timestart});
        }
    });
    res.send("Ok");
}

routeExp.route("/filter").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin"){
    session.filtrage = "";
  var searchit = req.body.searchit;
  var period = req.body.period;
  var datestart = "";
  var dateend = "";
  if (period == "t"){
    datestart = moment().format("YYYY-MM-DD");
  }
  else if (period == "tw"){
    datestart = moment().startOf("week").format("YYYY-MM-DD");
    dateend = moment().endOf("week").format("YYYY-MM-DD");
  }
  else if (period == "tm"){
    datestart = moment().startOf("month").format("YYYY-MM-DD");
    dateend = moment().endOf("month").format("YYYY-MM-DD");
  }
  else if  (period == "spec"){
    datestart = req.body.datestart;
    dateend= req.body.dateend
  }
  else{
    datestart = "";
    dateend = "";
  }
  var datecount = [];
  var datatosend = [];
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var late_temp =[];
      var absent_temp = [];
      var temp_leave = [];
      late_temp.push([]);
      absent_temp.push([]);
      temp_leave.push([]);
      datestart == "" ? "" : datecount.push(1);
      dateend == "" ? "" : datecount.push(2);
      searchit == "" ? delete session.request.search  : session.request.search = searchit;
      if (datecount.length == 2) {
        var day = moment
          .duration(
            moment(dateend, "YYYY-MM-DD").diff(moment(datestart, "YYYY-MM-DD"))
          )
          .asDays();
          var getdata;
          var getdatalate;
          var getdataabsence;
          var getdataleave;
        for (i = 0; i <= day; i++) {
          session.request.date = datestart;
          date_data.push(session.request.date);
         
          if (session.request.search){
            getdata = await StatusSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },
              { locaux: {'$regex':searchit,'$options' : 'i'}},],date:session.request.date,time_end:{ $ne: "" }}).sort({
              "_id": -1,
            });
            getdatalate = await LateSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },],date:session.request.date,validation:true}).sort({
              "_id": -1,
            });
            getdataabsence = await AbsentSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },],date:session.request.date,validation:true}).sort({
              "_id": -1,
            });
            getdataleave = await LeaveSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },],date_start:session.request.date}).sort({
              "_id": -1,
            });
          }
          else{
            getdata = await StatusSchema.find({date:session.request.date,time_end:{ $ne: "" }});
            getdatalate = await LateSchema.find({date:session.request.date,validation:true});
            getdataabsence = await AbsentSchema.find({date:session.request.date,validation:true});
            getdataleave = await LeaveSchema.find({date_start:session.request.date});
          }
          
          if (getdata.length != 0) {
            datatosend.push(getdata);
          }
          var addday = moment(datestart, "YYYY-MM-DD")
            .add(1, "days")
            .format("YYYY-MM-DD");
          datestart = addday;
          if (getdatalate.length != 0){
            for (l=0;l <getdatalate.length;l++){
              late_temp[0].push(getdatalate[l]);
          }
          }
          if (getdataabsence.length != 0){
            for (ab=0;ab <getdataabsence.length;ab++){
              absent_temp[0].push(getdataabsence[ab]);
            }
          }
          if (getdataleave.length != 0){
            for (lv =0;lv < getdataleave.length;lv++){
              temp_leave[0].push(getdataleave[lv]);
          }
          }
        }
        
        for (i = 1; i < datatosend.length; i++) {
         for (d=0;d<datatosend[i].length;d++){
            datatosend[0].push(datatosend[i][d]);
          }
        }
       
        if (datatosend.length != 0){
          session.datatowrite = datatosend[0];
          session.datalate = late_temp[0];
          session.dataabsence = absent_temp[0];
          session.dataleave = temp_leave[0];
          res.send(datatosend[0]);
        }
        else{
          session.datatowrite = datatosend;
          session.datalate = [];
          session.dataabsence = [];
          session.dataleave = [];
          res.send(datatosend);
        }
        
      } else if (datecount.length == 1) {
        if (datecount[0] == 1) {
          session.request.date = datestart;
          if (session.request.search){
            datatosend = await StatusSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },
              { locaux: {'$regex':searchit,'$options' : 'i'}},],date:session.request.date,time_end:{ $ne: "" }}).sort({
              "_id": -1,
            });
            session.datalate = await LateSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },],date:session.request.date,validation:true}).sort({
              "_id": -1,
            });
            session.dataabsence = await AbsentSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },],date:session.request.date,validation:true}).sort({
              "_id": -1,
            });
            session.dataleave = await LeaveSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },],date_start:session.request.date}).sort({
              "_id": -1,
            });
          }
          else{
            datatosend = await StatusSchema.find({date:session.request.date,time_end:{ $ne: "" }});
            session.datalate = await LateSchema.find({date:session.request.date,validation:true});
            session.dataabsence = await AbsentSchema.find({date:session.request.date,validation:true});
            session.dataleave = await LeaveSchema.find({date_start:session.request.date});
          }
          session.datatowrite = datatosend;
          session.searchit = searchit;
          res.send(datatosend);
        } else {
          session.request.date = dateend;
          if (session.request.search){
            datatosend = await StatusSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },
              { locaux: {'$regex':searchit,'$options' : 'i'}},],date:session.request.date,time_end:{ $ne: "" }}).sort({
              "_id": -1,
            });
            session.datalate = await LateSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },],date:session.request.date,validation:true}).sort({
              "_id": -1,
            });
            session.dataabsence = await AbsentSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },],date:session.request.date,validation:true}).sort({
              "_id": -1,
            });
            session.dataleave = await LeaveSchema.find({$or: 
              [{ m_code: {'$regex':searchit,'$options' : 'i'} },
              { nom: {'$regex':searchit,'$options' : 'i'} },],date_start:session.request.date}).sort({
              "_id": -1,
            });
          }
          else{
            datatosend = await StatusSchema.find({date:session.request.date});
          }
          session.datatowrite = datatosend;
          session.searchit = searchit;
          res.send(datatosend);
        }
      } else {
        delete session.request.date;
        datatosend = await StatusSchema.find({$or: 
          [{ m_code: {'$regex':searchit,'$options' : 'i'} },
          { nom: {'$regex':searchit,'$options' : 'i'} },
          { locaux: {'$regex':searchit,'$options' : 'i'}},],time_end:{ $ne: "" }}).sort({_id: -1,});
        session.datatowrite = datatosend;
        session.datalate = await LateSchema.find({$or: 
          [{ m_code: {'$regex':searchit,'$options' : 'i'} },
          { nom: {'$regex':searchit,'$options' : 'i'} },],validation:true});
        session.dataabsence = await AbsentSchema.find({$or: 
          [{ m_code: {'$regex':searchit,'$options' : 'i'} },
          { nom: {'$regex':searchit,'$options' : 'i'} },],validation:true});
        session.dataleave = await LeaveSchema.find({$or: 
          [{ m_code: {'$regex':searchit,'$options' : 'i'} },
          { nom: {'$regex':searchit,'$options' : 'i'} },]});
          
        session.searchit = searchit;
        res.send(datatosend);
      }
    });
  }
  else{
    res.send("error");
  }
});

routeExp.route("/leftwork").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User"){
    await leftwork(req.body.locaux,session,res);
  }
  else{
    res.send("error");
  }
});
// leftwork
async function leftwork(locaux,session,res){
  var date = moment().format("YYYY-MM-DD");
  var timeend = moment().add(3,'hours').format("HH:mm");
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
       await StatusSchema.findOneAndUpdate({m_code:session.m_code,date:date,locaux:locaux,time_end:""},{time_end:timeend});
       await UserSchema.findOneAndUpdate({m_code:session.m_code},{late:"n",count:0});
       session.occupation_u = null;
       res.redirect("/exit_u");
    });
}
//Page leave
routeExp.route("/leave").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin"){
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
    var alluser = await UserSchema.find({});
    res.render("conge.html",{users:alluser,notif:notification});
  })
}
else{
  res.redirect("/");
}
});
//List leave
routeExp.route("/leavelist").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var leavelist = await LeaveSchema.find({});
      res.render("congelist.html",{leavelist:leavelist,notif:notification});
    });
  }
  else{
  res.redirect("/");
  }
});
//take leave
routeExp.route("/takeleave").post(async function (req, res) {
    var userid = req.body.id;
    var type = req.body.type;
    var leavestart = req.body.leavestart;
    var leaveend = req.body.leaveend;
    mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
    var user = await UserSchema.findOne({_id:userid});
    if (await LeaveSchema.findOne({m_code:user.m_code,status:"on"})){
      res.send("already");
    }
    else{
    if (user.leave_stat == "y" && (type == "Paid Leave" || type == "Time off without pay")){
      var taked = date_diff(leavestart,leaveend);
      if (taked < user.remaining_leave){
    var new_leave = {
      m_code:user.m_code,
      num_agent : user.num_agent,
      nom:user.first_name + " "+ user.last_name,
      date_start:leavestart,
      date_end:leaveend,
      type:type,
      status:"on",
      validation:false
    }
    await LeaveSchema(new_leave).save();
    await UserSchema.findOneAndUpdate({_id:userid},{remaining_leave:(user.remaining_leave - taked),$inc:{leave_taked:taked},act_stat:"VACATION",act_loc:"Not defined"});
    const io = req.app.get('io');
    io.sockets.emit('status',"VACATION"+","+user.m_code);
    res.send("Ok");
  }
  else{
    res.send("exceeds");
  }
  }
  else if(type == "Sick Leave" || type == "Maternity / Paternity" ){
    var new_leave = {
      m_code:user.m_code,
      num_agent : user.num_agent,
      nom:user.first_name + " "+ user.last_name,
      date_start:leavestart,
      date_end:leaveend,
      type:type,
      status:"on",
      validation:false
    }
    await LeaveSchema(new_leave).save();
    await UserSchema.findOneAndUpdate({_id:userid},{act_stat:"VACATION",act_loc:"Not defined"});
    const io = req.app.get('io');
    io.sockets.emit('status',"VACATION"+","+user.m_code);
    res.send("Ok");
  }
  else{
    res.send("not authorized");
  }
}
  })
})
async function leave_permission(user){
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
        await UserSchema.findOneAndUpdate({m_code:user},{remaining_leave:30,leave_stat:"y"});
  })
}
//checkleave
async function checkleave(){
  mongoose
  .connect(
    "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
    {
      useUnifiedTopology: true,
      UseNewUrlParser: true,
    }
  )
  .then(async () => {
        var all_leave = await LeaveSchema.find({status:"on"});
        for (i=0;i< all_leave.length;i++){
            if (date_diff(moment().format("YYYY-MM-DD"),all_leave[i].date_end) <= 0){
              await UserSchema.findOneAndUpdate({m_code:all_leave[i].m_code},{act_stat:"LEFTING"});
              await LeaveSchema.findOneAndUpdate({_id:all_leave[i]._id},{status:"Done"});
              notification.push(all_leave[i].nom + " devrait revenir du cong??");
            }
        }
  })
}

routeExp.route("/statuschange").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User"){
    var locaux = req.body.act_loc;
      var status = req.body.act_stat;
      mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
       await UserSchema.findOneAndUpdate({m_code:session.m_code},{act_stat:status,act_loc:locaux});
       res.send(status);
    });

  }
  else{
    res.send("error");
  }
      
});
routeExp.route("/session_end").get(async function (req, res) {
    res.render("block.html");
})
//Generate excel file
routeExp.route("/generate").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
  var newsheet = ExcelFile.utils.book_new();
  newsheet.Props = {
    Title: "Timesheets",
    Subject: "Logged Time",
    Author: "Solumada",
  };
  newsheet.SheetNames.push("ALL USER");
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      
        var all_employes = [];
        for (i=0;i<session.datatowrite.length;i++){
           if (all_employes.includes(session.datatowrite[i].m_code)){
                
           }
           else{
             all_employes.push(session.datatowrite[i].m_code);
        }
        all_employes = all_employes.sort();
        }
        all_datas.push([
          "GLOBAL REPORT",
          "",
          "",
          "",
          "",
          "",
        ]);
        all_datas.push([
          "",
          "",
          "",
          "",
          "",
          "",
        ]);
        all_datas.push([
          "Name",
          "M-code",
          "Total time work",
          "Total delays",
          "Total absence",
          "Total leave work",
        ]);
        for (e = 0; e < all_employes.length; e++) {
          var name_user = await StatusSchema.findOne({m_code:all_employes[e]});
          data.push([
            "SHEET OF => "+name_user.nom,
            "",
            "",
            "",
            "",
            "",
          ]);
          data.push([
            "",
            "",
            "",
            "",
            "",
            "",
          ]);
          data.push([
            "M-code",
            "Number of agent",
            "Date",
            "Locaux",
            "Start Time",
            "End Time",
          ]);
            generate_excel(session.datatowrite,session.datalate,session.dataabsence,session.dataleave,all_employes[e]);
          if (newsheet.SheetNames.includes(all_employes[e])) {
          } else {
            newsheet.SheetNames.push(all_employes[e]);
          }
          newsheet.Sheets[all_employes[e]] = ws;
          hours = 0;
          minutes = 0;
          data = [];
        }
        global_Report(all_datas);
        newsheet.Sheets["ALL USER"] = ws;
        all_datas = [];
        if (newsheet.SheetNames.length != 0) {
            if (all_employes.length <= 1){
              session.filename = "N??"+num_file+" "+all_employes[0]+ ".xlsx";
              num_file++;
            }
            else{
              session.filename = "N??"+num_file+" Timesheets.xlsx";
              num_file++;
            }
            ExcelFile.writeFile(newsheet, session.filename);
          delete session.request.searchit;
          delete session.request.date;
          delete session.request.search;
          session.datatowrite = await StatusSchema.find();
        }
      res.send("Done");
    });
  }else{
    res.redirect("/");
  }
});
routeExp.route("/download").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    res.download(session.filename, function(err){
      fs.unlink(session.filename, function (err) {            
        if (err) {                                                 
            console.error(err);                                    
        }                                                          
       console.log('File has been Deleted');                           
    });         
    });
  }
});
//Add employee
routeExp.route("/addemp").post(async function (req, res) {
  session = req.session;
  var email = req.body.email;
  var mcode = req.body.mcode;
  var num_agent = req.body.num_agent;
  var change = "n";
  var first = req.body.first_name;
  var occupation = req.body.occupation;
  var last = req.body.last_name;
  var shift = req.body.shift;
  var late ="n";
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      if (
        await UserSchema.findOne({username: email})
      ) {
        res.send("error");
      } else {
        var passdefault = randomPassword();
        let hash = crypto.createHash('md5').update(passdefault).digest("hex");
        var new_emp = {
          username: email,
          first_name:first,
          last_name:last,
          password: hash,
          m_code: mcode,
          num_agent: num_agent,
          occupation: occupation,
          change:change,
          act_stat:"LEFTING",
          act_loc: "Not defined",
          shift:shift,
          late:late,
          count:0,
          take_break:"n",
          remaining_leave:0,
          leave_taked:0,
          leave_stat:"n",
          save_at:moment().format("YYYY-MM-DD"),
          user_ht:0
        };
        await UserSchema(new_emp).save();
        sendEmail(
          email,
          "Authentification Solumada",
          htmlRender(email, passdefault)
        );
        res.send(email);
      }
    });
});
//logout
routeExp.route("/exit_a").get(function (req, res) {
  session = req.session;
  session.occupation_a = null;
  res.redirect("/");
});
routeExp.route("/exit_u").get(function (req, res) {
  session = req.session;
  session.occupation_u = null;
  session.mcode = null;
  session.num_agent = null;

  res.redirect("/");
});
function htmlVerification(code) {
  return (
    "<center><h1>YOUR CLOCKING CODE AUTHENTIFICATION</h1>" +
    "<h3 style='width:250px;font-size:50px;padding:8px;background-color: rgba(87,184,70, 0.8); color:white'>" +
    code +
    "<h3></center>"
  );
}
function htmlRender(username, password) {
  var html =
    "<center><h1>Solumada Authentification</h1>" +
    '<table border="1" style="border-collapse:collapse;width:25%;border-color: lightgrey;">' +
    '<thead style="background-color: rgba(87,184,70, 0.8);color:white;font-weight:bold;height: 50px;">' +
    "<tr>" +
    '<td align="center">Username</td>' +
    '<td align="center">Password</td>' +
    "</tr>" +
    "</thead>" +
    '<tbody style="height: 50px;">' +
    "<tr>" +
    '<td align="center">' +
    username +
    "</td>" +
    '<td align="center">' +
    password +
    "</td>" +
    "</tr>" +
    "</tbody>" +
    "</table>";
  return html;
}
function randomPassword() {
  var code = "";
  let v = "abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ!??&#";
  for (let i = 0; i < 8; i++) {
    // 6 characters
    let char = v.charAt(Math.random() * v.length - 1);
    code += char;
  }
  return code;
}
function sendEmail(receiver, subject, text) {
  var mailOptions = {
    from: "Timesheets Optimum solution",
    to: receiver,
    subject: subject,
    html: text,
  };

  transporter.sendMail(mailOptions, function (error, info) {
    if (error) {
      console.log(error);
    } else {
      console.log("Email sent: " + info.response);
    }
  });
}
//Function Random code for verification
function randomCode() {
  var code = "";
  let v = "012345678";
  for (let i = 0; i < 6; i++) {
    // 6 characters
    let char = v.charAt(Math.random() * v.length - 1);
    code += char;
  }
  return code;
}
function calcul_timediff(startTime, endTime) {
  startTime = moment(startTime, "HH:mm:ss a");
  endTime = moment(endTime, "HH:mm:ss a");
  var duration = moment.duration(endTime.diff(startTime));
  //duration in hours
  hours += parseInt(duration.asHours());

  // duration in minutes
  minutes += parseInt(duration.asMinutes()) % 60;
  while (minutes > 60) {
    hours += 1;
    minutes = minutes - 60;
  }
  if (hours < 0 || minutes < 0){
    hours = hours + 24;
    if (minutes != 0){
      hours = hours - 1;
      minutes = minutes +60;
    }
  }
}
function hour_diff(startday,endday){
  startday = moment(startday, "HH:mm:ss a");
  endday = moment(endday, "HH:mm:ss a");
  var duration = moment.duration(endday.diff(startday));
  return parseInt(duration.asHours());
}
function convert_to_hour(mins){
  var hc = 0;
  while (mins > 60) {
    hc += 1;
    mins = mins - 60;
  }
  return hc +","+mins;
}
function difference_year(starting){
    var startings = moment(moment(starting)).format("YYYY-MM-DD");
    var nows = moment(moment().format("YYYY-MM-DD"),"YYYY-MM-DD");
    var duration = moment.duration(nows.diff(startings));
    var years = duration.years();
    return years;
}
function time_passed(starting){
  var startings = moment(moment(starting)).format("YYYY-MM-DD");
  var nows = moment(moment().format("YYYY-MM-DD"),"YYYY-MM-DD");
  var duration = moment.duration(nows.diff(startings));
  var years = duration.years();
  var months = duration.months();
  var days = duration.days();
  var tp = years +" years "+  months + " month(s) " + days + " day(s)";
  return tp;
}
function date_diff(starting,ending){
  var startings = moment(moment(starting)).format("YYYY-MM-DD");
  var endings = moment(ending,"YYYY-MM-DD");
  var duration = moment.duration(endings.diff(startings));
  var dayl = duration.days();
  return dayl;
}
function calcul_retard(regular,arrived){
  var time = 0;
  var lh = 0;
  var lm = 0;
  regular = moment(regular, "HH:mm:ss a");
  arrived = moment(arrived, "HH:mm:ss a");
  var duration = moment.duration(arrived.diff(regular));
  //duration in hours
  lh = parseInt(duration.asHours());
  // duration in minutes
  lm = parseInt(duration.asMinutes()) % 60;
  while (lm > 60) {
    lh += 1;
    lm = lm - 60;
  }
  lh = lh * 60;
  time = lh + lm;
  return time;
}
function style(){
  var cellule = ["A", "B", "C", "D", "E", "F","G"];
  for (c = 0; c < cellule.length; c++) {
    for (i = 1; i <= data.length; i++) {
      if (ws[cellule[c] + "" + i]) {
        if (i == 1 || i == 2) {
          ws[cellule[c] + "" + i].s = {
            font: {
              name: "Segoe UI Black",
              bold: true,
              color: { rgb: "398C39" },
            },
            alignment:{
                vertical : "center",
                horizontal:"center"
            },
          };
        }
        else if (i == 3) {
          ws[cellule[c] + "" + i].s = {
            fill:{
              patternType : "solid",
              fgColor : { rgb: "398C39" },
              bgColor: { rgb: "398C39" },
            },
            font: {
              name: "Segoe UI Black",
              bold: true,
              color: { rgb: "F5F5F5" }
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment:{
                vertical : "center",
                horizontal:"center"
            },
          };
        } 
        else {
          ws[cellule[c] + "" + i].s = {
            font: {
              name: "Verdana",
              color: {rgb:"777777"}
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment:{
              vertical : "center",
              horizontal:"center"
          },
          };
        }
      }
    }
  }
}
function style2(){
  var cellule = ["A", "B", "C", "D", "E", "F","G"];
  for (c = 0; c < cellule.length; c++) {
    for (i = 1; i <= all_datas.length; i++) {
      if (ws[cellule[c] + "" + i]) {
        if (i == 1 || i == 2) {
          ws[cellule[c] + "" + i].s = {
            font: {
              name: "Segoe UI Black",
              bold: true,
              color: { rgb: "398C39" },
            },
            alignment:{
                vertical : "center",
                horizontal:"center"
            },
          };
        }
        else if (i == 3) {
          ws[cellule[c] + "" + i].s = {
            fill:{
              patternType : "solid",
              fgColor : { rgb: "398C39" },
              bgColor: { rgb: "398C39" },
            },
            font: {
              name: "Segoe UI Black",
              bold: true,
              color: { rgb: "F5F5F5" }
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment:{
                vertical : "center",
                horizontal:"center"
            },
          };
        } 
        else {
          ws[cellule[c] + "" + i].s = {
            font: {
              name: "Verdana",
              color: {rgb:"777777"}
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment:{
              vertical : "center",
              horizontal:"center"
          },
          };
        }
      }
    }
  }
}
//generate excel
function readfile(name_file){
  // Requiring the module
  
// Reading our test file
const file = ExcelFile.readFile(name_file)
  
let data = []
  
const sheets = file.SheetNames
  
for(let i = 0; i < sheets.length; i++)
{
   const temp = ExcelFile.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
   temp.forEach((res) => {
      data.push(res)
   })
}
  
// Printing data
console.log(data);
}
//Fonction generate excel
function generate_excel(datatowrites,retard,absent,conge,code) {
  var counter = 0;
  var cum = 0;
  var cumg = 0;
  var cum_tot = "";
  var cum_abs ="";
  var cum_del = "";
  var nom = "";
  var m_codes = "";
  for (i = 0; i < datatowrites.length; i++) { 
    if (datatowrites[i].time_end != "" && datatowrites[i].m_code == code){
      counter++;
      var ligne = [
        datatowrites[i].m_code,
        datatowrites[i].num_agent,
        datatowrites[i].date,
        datatowrites[i].locaux,
        datatowrites[i].time_start,
        datatowrites[i].time_end,
      ];
      nom = datatowrites[i].nom;
      m_codes = datatowrites[i].m_code;
      data.push(ligne);
      calcul_timediff(datatowrites[i].time_start, datatowrites[i].time_end);
    }
  }
  totaltime = hours + "H " + minutes + "MN";
  data.push(["", "", "", "", "TOTAL",totaltime]);
  cum_tot = totaltime;
  data.push(["", "", "", "", "",""]);
  hours = 0;minutes =0;
  if(retard.length != 0){
    if (retard[0].m_code == code){
    data.push(["", "", "", "", "",""]);
    data.push(["", "", "", "DELAYS REPORT", "",""]);
    data.push(["M-code", "Number of agent", "Date", "Reason", "Time",""]);
    for ( i = 0;i<retard.length ; i++){
      if (retard[i].m_code == code){
      cum+=retard[i].time;
        var lateligne = [
          retard[i].m_code,
          retard[i].num_agent,
          retard[i].date,
          retard[i].reason,
          retard[i].time + " minutes",
        ];
        data.push(lateligne);
      }
    }
    if (cum != 0){
      cum = convert_to_hour(cum).split(",");
      cum_del =  cum[0] + "H" + " "+cum[1] + " MN";
      data.push(["", "", "","", "TOTAL", cum[0] + "H" + " "+cum[1] + " MN"]);
    }
    else{
      data.push(["", "", "","", "TOTAL","0H" + " "+"0 MN"]);
    }
   
  }
  }
  if(absent.length != 0){
    if(absent[0].m_code == code){
    data.push(["", "", "", "", "",""]);
    data.push(["", "", "", "ABSENCE REPORT (RETURN)", "","",""]);
    data.push(["M-code", "Number of agent", "Date", "Reason", "Time Start","Return","Status"]);
    for ( i = 0;i<absent.length ; i++){
      var latelignent = [];
      if (absent[i].return != "Not come back" && absent[i].m_code == code){
        var lateligne = [
          absent[i].m_code,
          absent[i].num_agent,
          absent[i].date,
          absent[i].reason,
          absent[i].time_start,
          absent[i].return,
          absent[i].status,
        ];
        data.push(lateligne);
        calcul_timediff(absent[i].time_start, absent[i].return);
      }
      else{
        if (absent[i].m_code == code){
        latelignent.push(
          absent[i].m_code,
          absent[i].num_agent,
          absent[i].date,
          absent[i].reason,
          absent[i].time_start,
          absent[i].return,
          absent[i].status,
        );
        }
      }
        
    }
    totaltime = hours + "H " + minutes + "MN";
    cum_abs = totaltime;
    data.push(["", "", "", "","", "TOTAL",totaltime]);
    data.push(["", "", "", "ABSENCE REPORT ( NO RETURN)", "","",""]);
    data.push(["M-code", "Number of agent", "Date", "Reason", "Time Start","Return","Status"]);
    data.push(latelignent);
  }
    
  }
  if(conge.length != 0){
    if(conge[0].m_code == code){
    data.push(["", "", "", "", ""]);
    data.push(["", "", "", "LEAVE FROM WORK", ""]);
    data.push(["M-code", "Number of agent", "Date Start", "Date end", "Type"]);
    for ( i = 0;i<conge.length ; i++){
      if (conge[i].m_code == code){
        var lateligne = [
          conge[i].m_code,
          conge[i].num_agent,
          conge[i].date_start,
          conge[i].date_end,
          conge[i].type,
        ];
        data.push(lateligne);
        cumg += date_diff(conge[i].date_start,conge[i].date_end);
    }
  }
    data.push(["", "", "", "TOTAL", cumg + "days"]);
  }
  }
  each_data = [
    nom,
    m_codes,
    cum_tot,
    cum_del,
    cum_abs,
    cumg + " day(s)"
  ]
  all_datas.push(each_data);
  ws = ExcelFile.utils.aoa_to_sheet(data);
  ws["!cols"] = [
    {wpx: 80 },
    {wpx: 130},
    {wpx: 200},
    {wpx: 250},
    {wpx: 160},
    {wpx: 100 },
    {wpx: 100 }
  ];
  const merge = [
    { s: { r: 0, c: 0 }, e: { r: 1, c: 5}},
    { s: { r: 3, c: 0 }, e: { r: counter + 2, c:0 }},
    { s: { r: 3, c: 1 }, e: { r: counter + 2, c:1 }},
    {s: {r:counter +5,c:0},e: {r:counter+5,c:5}}
  ];
  ws["!merges"] = merge;
  style();
}
function global_Report(all_data){
  ws = ExcelFile.utils.aoa_to_sheet(all_data);
  ws["!cols"] = [
    {wpx: 230 },
    {wpx: 80},
    {wpx: 150},
    {wpx: 150},
    {wpx: 150},
    {wpx: 150 },
    {wpx: 150 }
  ];
  const merge = [
    { s: { r: 0, c: 0 }, e: { r: 1, c: 5}}
  ];
  ws["!merges"] = merge;
  style2();
}

module.exports = routeExp;
