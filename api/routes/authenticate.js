const express = require('express'); //to create instance of express module
const mysql   = require('mysql'); //to create mysql instance
const loadScript = Symbol('loadScript');
const router  = express.Router();

//include DbHandler
const DBHandler = require("../function/DbHandler");
var DbHandler   = new DBHandler(); 

//sql DbConnection
const dbConnection  = mysql.createConnection({
    //DbProperties
    host : 'localhost',
    user : 'root',
    password  : '',
    database : 'user'
});

//dbconnection response message
dbConnection.connect(function(error) {
    if(!!error){
        console.log('Error');
    }else{
            console.log('Connected');
    }
});


/*************
 **CREATE API*
 *************/
router.post('/registeruser',(req, res, next) => {
    var username    =    req.body.username;
    var email       =    req.body.email;
    var password    =    req.body.password;

    var checkuser   =    DbHandler.checkUser(username,email).then(function(value) { 
        console.log(value);
        if(value == 0) {
        var insertUser  =    DbHandler.insertUser(username,email,password).then(function(insertResponse) {;
            console.log(insertResponse);    
            if(insertResponse == 0) {
                res.status(201).json({
                    error   : false,
                    message : "User Successfully Registered"    
             });
            } else if(insertResponse == 1) {
                res.status(201).json({
                    error   : true,
                    message : "User Registration Failed"    
              });
            } 
         }); //end of insertUser      
        } else  if(value == 1) {
            res.status(201).json({
                     error   : true,
                     message : "User with same credentials exist's"    
            });
        }//end of else
    });//end of checkuser        
});


/************
 *LOGIN API*
 ************/
router.post('/login',(req, res, next) => {
    var email       =    req.body.email;
    var password    =    req.body.password;

    var checkuser   =  DbHandler.login(email,password).then(function(value) {
        if(value == 0) {
        res.status(201).json({
            error   : false,
            message : "User Login Successfull"    
        });
        } else if(value == 1) {
        res.status(201).json({
            error   : true,
            message : "email or password mismatch"    
         });
        }
    })     
});

/**********************
 *Forget and Reset pwd* 
 *********************/
 router.post('/forgotpassword',(req,res,next) => {

    //post parameters
    var email       =    req.body.email;
    var newpassword =    req.body.newpassword;
    var statusMsg,userid;

    //sql query
    dbConnection.query("SELECT * FROM users WHERE email = '"+email+"'",function(error,result,fields){
        if(!!error){
                    return error;
        } else {
                // console.log(result);
                if(result != null && result.length != 0){
                    statusMsg    =   0; // user exist's
                    userid       = result[0].userid;    
                 } else {
                    statusMsg    =   1; // user does not exist's
                }
       }//end of else
    });//end of query execution function
    
    setTimeout(function () {
        if(statusMsg == 0) {
            console.log(userid);
            dbConnection.query("UPDATE users SET password = '"+newpassword+"' WHERE userid = '"+userid+"'",function(error,result,fields){
                if(!!error){
                            return error;
                } else {
                        if(result != null && result.affectedRows > 0){
                            res.status(201).json({
                                error   : true,
                                message : "User pwd as been updated succeesfully"    
                           });
                        } else {
                            res.status(201).json({
                                error   : true,
                                message : "Updation Failed"    
                           });
                       }
               }//end of else
            });//end of query execution function        
        } else {
            res.status(201).json({
                 error   : true,
                 message : "User with mentioned credentials does not exist's"    
            });
        }//end of else
    },100);
 });


module.exports = router;