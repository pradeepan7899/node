const express = require('express'); //to create instance of express module
const router  = express.Router();

const mysql   = require('mysql'); //to create mysql instance

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


class DbHandler{

    checkUser(username,email) {
        console.log(username,email);
      return new Promise(function(resolve, reject) {
         //sql query
         dbConnection.query("SELECT * FROM users WHERE username = '"+username+"' AND email = '"+email+"'",function(error,result,fields){
            if(!!error){
                console.log(error);
         } else {
             if(result.length == 0){
                resolve(0);// success response or user deos not exists confirmation
            } else {
                resolve(1);//user exist's    
            }
         }//end of else
        });//end of query execution function
      })//end of promise   
    } // end of checkuser function

    insertUser(username,email,password)
    {
        //new promise
        return new Promise(function(resolve, reject) {
         //insert query   
        dbConnection.query("INSERT INTO users(userid,username,email,password) values('','"+username+"','"+email+"','"+password+"')",function(error,result,fields){

        if(!!error) {
               console.log(error); 
        } else {
                    if(result.affectedRows == 1) {
                        resolve(0);// data insertion successfully
                    } else {
                        resolve(1);//data insertion failed
                    }
        }//end of else      
       });//end of dbConnection
      });//end of promise
    }

    login(email,password) {
        //promise
        return new Promise(function(resolve, reject) {
            //sql query
            dbConnection.query("SELECT * FROM users WHERE email = '"+email+"' AND password = '"+password+"'",function(error,result,fields){
                     if(!!error){
                         console.log(error);
                  } else {
                     console.log(result);
                     if(result != null && result.length != 0){
                         resolve(0);// success response or user deos not exists confirmation
                     } else {
                        resolve(1);    
                     }
                  }//end of else
        });//end of query execution function
      })//end of promise   
    }

}

module.exports = DbHandler;