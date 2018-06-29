const express = require('express'); //to create instance of express module
const mysql   = require('mysql'); //to create mysql instance
const loadScript = Symbol('loadScript');
const router  = express.Router();

//used for pwd encrypting / decrypting purpouse
const bcrypt  = require('bcrypt');

//excel-class
const Excel = require('excel-class');
// const Excel = require('exceljs');

//s3 bucket
const AWS = require('aws-sdk');
const fs = require('fs');
const path = require('path');

//firebase
const firebase = require("firebase"); //for creating firebase instance

//need to add firebase database url
/*To get databaseURL, 
 *Go to the Firebase Console -> DEVELOP → Database → DATA Tab → https://yourdatabaseURL.com/
 */
firebase.initializeApp({   
    databaseURL : "https://shrirampropertyconstruction.firebaseio.com/"
}); 

//sql DbConnection
// const dbConnection  = mysql.createConnection({
//     //DbProperties
//     host : 'localhost',
//     user : 'root',
//     password  : '',
//     database : 'propertydb'
// });

// dbConnection.connect(function(error) {
//         if(!!error){
//             console.log('Error');
//         }else{
//                 console.log('Connected');
//         }
// });

//api for common lift activity
router.post('/liftactivities', function(req,res) {
    const dbRef = firebase.database().ref("commonlift_activities");
  
    //Panorama Hills
    var arr = new Array();
    var arr2 = new Array();
    var arr3 = new Array();
    var arr4 = new Array();
    var arr5 = new Array();
  
    //Greenfield 1
    var arr6 = new Array();
    var arr7 = new Array();
    var arr8 = new Array();
    var arr9 = new Array();
    var arr10 = new Array();
    var arr11 = new Array();
  
    //Greenfield 2
    var arr12 =  new Array();
    var arr13 =  new Array();
    var arr14 =  new Array();
  
    //Southern Crest
    var arr15 = new Array();
    var arr16 = new Array();
    var arr17 = new Array();
  
    //Luxor
    var arr18 = new Array();
    var arr19 = new Array();
    var arr20 = new Array();
    var arr21 = new Array();
    var arr22 = new Array();
    var arr23 = new Array();
    var arr24 = new Array();
    var arr25 = new Array();
  
    var obj1 = {};
    var obj2 = {};
    var obj3 = {};
    var obj4 = {};
  
  
  
    dbRef.once("value").then(snap => {
  
      //Panorama Hills
        var T2ACount = 0;
        var T2BCount = 0;
        var T2CCount = 0;
        var T2DCount = 0;
  
      //Greenfield 1
        var TowerGA = 0;
        var TowerGB = 0;
        var TowerGC = 0;
        var TowerGD = 0;
        var TowerGE = 0;
        var TowerGF = 0;
  
      //Greenfield 2
        var TowerGG = 0;
        var TowerGH = 0;
        var TowerGJ = 0;
  
      //Southern Crest
        var TowerSA = 0;
        var TowerSB = 0;
        var TowerSC = 0;
  
      //Luxor
        var TowerLA = 0;
        var TowerLB = 0;
        var TowerLC = 0;
        var TowerLD = 0;
        var TowerLE = 0;
        var TowerLF = 0;
        var TowerLG = 0;
        var TowerLH = 0;
  
        var liftActiviytyValue = snap.val();
  
        //loop for getting total number of tower in the flat
        for(var i=1;i<liftActiviytyValue.length;i++)
        {
  
          //Panorama Hills
          if(liftActiviytyValue[i].project_name == "Panorama Hills")
          {   
            if(liftActiviytyValue[i].tower == "2A" && T2ACount == 0)
            {
                T2ACount++;
                obj1['tower_name'] = "Tower 2A";
                obj1['total_unit'] = T2ACount;
                arr.push(obj1);
            }
            else if(liftActiviytyValue[i].tower == "2B" && T2BCount == 0)
            {
                T2BCount++;
                obj2['tower_name'] = "Tower 2B";
                obj2['total_unit'] = T2BCount;
                arr.push(obj2);   
            }
            else if(liftActiviytyValue[i].tower == "2C" && T2CCount == 0)
            {
                T2CCount++;
                obj3['tower_name'] = "Tower 2C";
                obj3['total_unit'] = T2CCount;
                arr.push(obj3);  
            }
            else if(liftActiviytyValue[i].tower == "2D" && T2DCount == 0)
            {
                T2DCount++;
                obj4['tower_name'] = "Tower 2D";
                obj4['total_unit'] = T2DCount; 
                arr.push(obj4);  
            }
          }
  
          //Greenfield 1
          else if(liftActiviytyValue[i].project_name == "Greenfield Phase 1")
            {
                var obj = {};
  
              if(liftActiviytyValue[i].tower == "Tower A" && TowerGA == 0)
              {
                TowerGA++;
                obj['tower_name'] = "Tower A";
                obj['total_unit'] = TowerGA;
                arr.push(obj);
              }
              else if(liftActiviytyValue[i].tower == "Tower B" && TowerGB == 0)
              {
                  TowerGB++;
                  obj['tower_name'] = "Tower B";
                  obj['total_unit'] = TowerGB;
                  arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower C" && TowerGC == 0)
              {
                  TowerGC++;
                  obj['tower_name'] = "Tower C";
                  obj['total_unit'] = TowerGC;
                  arr.push(obj);                                  
              }
              else if(liftActiviytyValue[i].tower == "Tower D" && TowerGD == 0)
              {
                  TowerGD++;
                  obj['tower_name'] = "Tower D";
                  obj['total_unit'] = TowerGD;
                  arr.push(obj);                                  
              }
              else if(liftActiviytyValue[i].tower == "Tower E" && TowerGE == 0)
              {
                  TowerGE++;
                  obj['tower_name'] = "Tower E";
                  obj['total_unit'] = TowerGE;
                  arr.push(obj);                                  
              }
              else if(liftActiviytyValue[i].tower == "Tower F" && TowerGF == 0)
              {
                  TowerGF++;
                  obj['tower_name'] = "Tower F";
                  obj['total_unit'] = TowerGF;
                  arr.push(obj);                                  
              }
            }
  
          //Greenfield 2
          else if(liftActiviytyValue[i].project_name == "Greenfield Phase 2")
          {
  
              var obj = {};
              if(liftActiviytyValue[i].tower == "Tower G" && TowerGG == 0)
              {
                   TowerGG++;
                   obj['tower_name'] = "Tower G";
                   obj['total_unit'] = TowerGG;
                   arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower H" && TowerGH == 0)
              {
                TowerGH++;
                obj['tower_name'] = "Tower H";
                obj['total_unit'] = TowerGH;
                arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower J" && TowerGJ == 0)
              {
                TowerGJ++;
                obj['tower_name'] = "Tower J";
                obj['total_unit'] = TowerGJ;
                arr.push(obj);                  
              }
          }
  
        //Southern Crest
          else if(liftActiviytyValue[i].project_name == "Southern Crest")
            {
              var obj = {};
              if(liftActiviytyValue[i].tower == "Tower A" && TowerSA == 0)
              {
                TowerSA++;
                obj['tower_name'] = "Tower A";
                obj['total_unit'] = TowerSA;
                arr.push(obj);
              }
              else if(liftActiviytyValue[i].tower == "Tower B" && TowerSB == 0)
              {
                  TowerSB++;
                  obj['tower_name'] = "Tower B";
                  obj['total_unit'] = TowerSB;
                  arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower C" && TowerSC == 0)
              {
                  TowerSC++;
                  obj['tower_name'] = "Tower C";
                  obj['total_unit'] = TowerSC;
                  arr.push(obj);                  
              }
            }
  
        //Luxor
          else if(liftActiviytyValue[i].project_name == "Luxor")
            {
              var obj = {};
              if(liftActiviytyValue[i].tower == "Tower A" && TowerLA == 0)
              {
                TowerLA++;
                obj['tower_name'] = "Tower A";
                obj['total_unit'] = TowerLA;
                arr.push(obj);
              }
              else if(liftActiviytyValue[i].tower == "Tower B" && TowerLB == 0)
              {
                TowerLB++;
                obj['tower_name'] = "Tower B";
                obj['total_unit'] = TowerLB;
                arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower C" && TowerLC == 0)
              {
                TowerLC++;
                obj['tower_name'] = "Tower C";
                obj['total_unit'] = TowerLC;
                arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower D" && TowerLD == 0)
              {
                TowerLD++;
                obj['tower_name'] = "Tower D";
                obj['total_unit'] = TowerLD;
                arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower E" && TowerLE == 0)
              {
                TowerLE++;
                obj['tower_name'] = "Tower E";
                obj['total_unit'] = TowerLE;
                arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower F" && TowerLF == 0)
              {
                TowerLF++;
                obj['tower_name'] = "Tower F";
                obj['total_unit'] = TowerLF;
                arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower G" && TowerLG == 0)
              {
                TowerLG++;
                obj['tower_name'] = "Tower G";
                obj['total_unit'] = TowerLG;
                arr.push(obj);                  
              }
              else if(liftActiviytyValue[i].tower == "Tower H" && TowerLH == 0)
              {
                TowerLH++;
                obj['tower_name'] = "Tower H";
                obj['total_unit'] = TowerLH;
                arr.push(obj);                  
              }
            }
        }
  
    
  
        //loop to checkmilestone status
        for(var j=1;j<liftActiviytyValue.length;j++)
        {
            // milestone activities
            var comp = {};
            var uncomp = {};
  
        //Panorama Hills
          if(liftActiviytyValue[j].project_name == "Panorama Hills" && liftActiviytyValue[j].tower == "2A")
          {  
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr2.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr2.push(uncomp);        
              }
          }      
          else if(liftActiviytyValue[j].project_name == "Panorama Hills" && liftActiviytyValue[j].tower == "2B")
          {   
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr3.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr3.push(uncomp);        
              }
            }
            else if(liftActiviytyValue[j].project_name == "Panorama Hills" && liftActiviytyValue[j].tower == "2C")
            {   
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr4.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr4.push(uncomp);        
              }
            }    
            else if(liftActiviytyValue[j].project_name == "Panorama Hills" && liftActiviytyValue[j].tower == "2D")
            {   
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr5.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr5.push(uncomp);        
              }
            }
  
          //Greenfield 1
            else if(liftActiviytyValue[j].project_name == "Greenfield Phase 1" && liftActiviytyValue[j].tower == "Tower A")
            {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr6.push(comp);
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr6.push(uncomp);        
              }  
            }
            else if(liftActiviytyValue[j].project_name == "Greenfield Phase 1" && liftActiviytyValue[j].tower == "Tower B")
            {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr7.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr7.push(uncomp);        
              }  
            }
            else if(liftActiviytyValue[j].project_name == "Greenfield Phase 1" && liftActiviytyValue[j].tower == "Tower C")
            {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr8.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr8.push(uncomp);        
              }  
            }
            else if(liftActiviytyValue[j].project_name == "Greenfield Phase 1" && liftActiviytyValue[j].tower == "Tower D")
            {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr9.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr9.push(uncomp);        
              }  
            }
            else if(liftActiviytyValue[j].project_name == "Greenfield Phase 1" && liftActiviytyValue[j].tower == "Tower E")
            {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr10.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr10.push(uncomp);        
              }  
            }
            else if(liftActiviytyValue[j].project_name == "Greenfield Phase 1" && liftActiviytyValue[j].tower == "Tower F")
            {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr11.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr11.push(uncomp);        
              }  
            }
  
          //Greenfield 2
            else if(liftActiviytyValue[j].project_name == "Greenfield Phase 2" && liftActiviytyValue[j].tower == "Tower G")
              {
                  if(liftActiviytyValue[j].activity_status == "1")
                  {
                    var completed = 1;
                    comp['milestone'] = liftActiviytyValue[j].activity_name;
                    comp['count'] = completed;
                    arr12.push(comp);            
                  }
                  else if(liftActiviytyValue[j].activity_status == "0")
                  {
                    var uncompleted = 0;
                    uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                    uncomp['count'] = uncompleted;
                    arr12.push(uncomp);        
                  }   
              }
              else if(liftActiviytyValue[j].project_name == "Greenfield Phase 2" && liftActiviytyValue[j].tower == "Tower H")
              {
                  if(liftActiviytyValue[j].activity_status == "1")
                  {
                    var completed = 1;
                    comp['milestone'] = liftActiviytyValue[j].activity_name;
                    comp['count'] = completed;
                    arr13.push(comp);            
                  }
                  else if(liftActiviytyValue[j].activity_status == "0")
                  {
                    var uncompleted = 0;
                    uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                    uncomp['count'] = uncompleted;
                    arr13.push(uncomp);        
                  }   
              }
              else if(liftActiviytyValue[j].project_name == "Greenfield Phase 2" && liftActiviytyValue[j].tower == "Tower J")
              {
                  if(liftActiviytyValue[j].activity_status == "1")
                  {
                    var completed = 1;
                    comp['milestone'] = liftActiviytyValue[j].activity_name;
                    comp['count'] = completed;
                    arr14.push(comp);            
                  }
                  else if(liftActiviytyValue[j].activity_status == "0")
                  {
                    var uncompleted = 0;
                    uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                    uncomp['count'] = uncompleted;
                    arr14.push(uncomp);        
                  }   
              }
  
              //Southern Crest
              else if(liftActiviytyValue[j].project_name == "Southern Crest" && liftActiviytyValue[j].tower == "Tower A")
              {
                  if(liftActiviytyValue[j].activity_status == "1")
                  {
                    var completed = 1;
                    comp['milestone'] = liftActiviytyValue[j].activity_name;
                    comp['count'] = completed;
                    arr15.push(comp);            
                  }
                  else if(liftActiviytyValue[j].activity_status == "0")
                  {
                    var uncompleted = 0;
                    uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                    uncomp['count'] = uncompleted;
                    arr15.push(uncomp);        
                  }  
              }
              else if(liftActiviytyValue[j].project_name == "Southern Crest" && liftActiviytyValue[j].tower == "Tower B")
              {
                if(liftActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = liftActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr16.push(comp);            
                }
                else if(liftActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr16.push(uncomp);        
                }              
              }
              else if(liftActiviytyValue[j].project_name == "Southern Crest" && liftActiviytyValue[j].tower == "Tower C")
              {
                if(liftActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = liftActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr17.push(comp);            
                }
                else if(liftActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr17.push(uncomp);        
                }
              }
  
          //Luxor
          else if(liftActiviytyValue[j].project_name == "Luxor" && liftActiviytyValue[j].tower == "Tower A")
          {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr18.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr18.push(uncomp);        
              }   
          }
          else if(liftActiviytyValue[j].project_name == "Luxor" && liftActiviytyValue[j].tower == "Tower B")
          {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr19.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr19.push(uncomp);        
              }   
          }    
          else if(liftActiviytyValue[j].project_name == "Luxor" && liftActiviytyValue[j].tower == "Tower C")
          {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr20.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr20.push(uncomp);        
              }   
          }
          else if(liftActiviytyValue[j].project_name == "Luxor" && liftActiviytyValue[j].tower == "Tower D")
          {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr21.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr21.push(uncomp);        
              }   
          }    
          else if(liftActiviytyValue[j].project_name == "Luxor" && liftActiviytyValue[j].tower == "Tower E")
          {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr22.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr22.push(uncomp);        
              }   
          }
          else if(liftActiviytyValue[j].project_name == "Luxor" && liftActiviytyValue[j].tower == "Tower F")
          {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr23.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr23.push(uncomp);        
              }   
          }
          else if(liftActiviytyValue[j].project_name == "Luxor" && liftActiviytyValue[j].tower == "Tower G")
          {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr24.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr24.push(uncomp);        
              }   
          }
          else if(liftActiviytyValue[j].project_name == "Luxor" && liftActiviytyValue[j].tower == "Tower H")
          {
              if(liftActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = liftActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr25.push(comp);            
              }
              else if(liftActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = liftActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr25.push(uncomp);        
              }   
          }
          }        
  
      //Generate Excel for Panorama Hills
        generateCommonLiftExcelVizag(arr,arr2,arr3,arr4,arr5);
      //Generate Excel for Greenfield Phase1
        generateCommonLiftExcelGFPhase1(arr,arr6,arr7,arr8,arr9,arr10,arr11);
      //Generate Excel for Greenfield Phase2
        generateCommonLiftExcelGFPhase2(arr,arr12,arr13,arr14);
      //Generate Excel for Southern Crest
        generateCommonLiftExcelSC(arr,arr15,arr16,arr17);
      //GenerateExcel for Luxor
        generateCommonLiftExcelLuxor(arr,arr18,arr19,arr20,arr21,arr22,arr23,arr24,arr25);
  
        res.status(200).send({
            error:false,
            Array:"Exported to Excel Sheet",
        });
  
    }).catch(error => {
        console.log("error".error)
    });
  });


//api for other activities
router.post('/other_activities',(req,res,next) =>{
const dbRef = firebase.database().ref("commonother_activities");

var arr  = new Array();


    //arr for greenphase-1 
    var arr6 = new Array();
    var arr7 = new Array();
    var arr8 = new Array();
    var arr9 = new Array();
    var arr10 = new Array();
    var arr11 = new Array();

    //arr for Southren crest 
    var arr12 = new Array();
    var arr13 = new Array();
    var arr14 = new Array();

    //arr for Luxor
    var arr15 = new Array();
    var arr16 = new Array();
    var arr17 = new Array();
    var arr18 = new Array();
    var arr19 = new Array();
    var arr20 = new Array();
    var arr21 = new Array();
    var arr22 = new Array();

    //arr for greenfield phase2
    var arr23 =  new Array();
    var arr24 =  new Array();
    var arr25 =  new Array();


var Tower2A = new Array();
var Tower2B = new Array();
var Tower2C = new Array();
var Tower2D = new Array();

  dbRef.once("value").then(snap => {
      //panorama hills
      var T2ACount = 0;
      var T2BCount = 0;
      var T2CCount = 0;
      var T2DCount = 0;

       //greenfield pahse1
       var TowerA = 0;
       var TowerB = 0;
       var TowerC = 0;
       var TowerD = 0;
       var TowerE = 0;
       var TowerF = 0;
      
       //greenfield pahse2
       var TowerG = 0;
       var TowerH = 0;
       var TowerJ = 0;
      
       //southrencrest variable
       var TowerSA = 0;
       var TowerSB = 0;
       var TowerSC = 0;
      
       //LUXOR Varirables
       var TowerLA = 0;
       var TowerLB = 0;
       var TowerLC = 0;
       var TowerLD = 0;
       var TowerLE = 0;
       var TowerLF = 0;
       var TowerLG = 0;
       var TowerLH = 0;
      
      var otherActiviytyValue = snap.val();

      for(var i=1;i<otherActiviytyValue.length;i++)
      {
        if(otherActiviytyValue[i].project_name == "Panorama Hills")
        {   
          var obj = {};

          if(otherActiviytyValue[i].tower == "2A" && T2ACount == 0)
          {
              T2ACount++;
              obj['project_name'] = "Panorama Hills";
              obj['tower_name'] = "Tower 2A";
              obj['total_unit'] = T2ACount;
              arr.push(obj);
          }
          else if(otherActiviytyValue[i].tower == "2B" && T2BCount == 0)
          {
              T2BCount++;
              obj['tower_name'] = "Tower 2B";
              obj['total_unit'] = T2BCount;
              arr.push(obj);   
          }
          else if(otherActiviytyValue[i].tower == "2C" && T2CCount == 0)
          {
              T2CCount++;
              obj['tower_name'] = "Tower 2C";
              obj['total_unit'] = T2CCount;
              arr.push(obj);  
          }
          else if(otherActiviytyValue[i].tower == "2D" && T2DCount == 0)
          {
              T2DCount++;
              obj['tower_name'] = "Tower 2D";
              obj['total_unit'] = T2DCount; 
              arr.push(obj);  
          }
        }
        else if(otherActiviytyValue[i].project_name == "Greenfield Phase 1")
        {
            var obj = {};

          if(otherActiviytyValue[i].tower == "Tower A" && TowerA == 0)
          {
            TowerA++;
            obj['project_name'] = "Greenfield Phase 1";
            obj['tower_name'] = "Tower A";
            obj['total_unit'] = TowerA;
            arr.push(obj);
          }
          else if(otherActiviytyValue[i].tower == "Tower B" && TowerB == 0)
          {
              TowerB++;
              obj['tower_name'] = "Tower B";
              obj['total_unit'] = TowerB;
              arr.push(obj);                  
          }
          else if(otherActiviytyValue[i].tower == "Tower C" && TowerC == 0)
          {
              TowerC++;
              obj['tower_name'] = "Tower C";
              obj['total_unit'] = TowerC;
              arr.push(obj);                                  
          }
          else if(otherActiviytyValue[i].tower == "Tower D" && TowerD == 0)
          {
              TowerD++;
              obj['tower_name'] = "Tower D";
              obj['total_unit'] = TowerD;
              arr.push(obj);                                  
          }
          else if(otherActiviytyValue[i].tower == "Tower E" && TowerE == 0)
          {
              TowerE++;
              obj['tower_name'] = "Tower E";
              obj['total_unit'] = TowerE;
              arr.push(obj);                                  
          }
          else if(otherActiviytyValue[i].tower == "Tower F" && TowerF == 0)
          {
              TowerF++;
              obj['tower_name'] = "Tower F";
              obj['total_unit'] = TowerF;
              arr.push(obj);                                  
          }
        }
        else if(otherActiviytyValue[i].project_name == "Greenfield Phase 2")
        {

            var obj = {};
            if(otherActiviytyValue[i].tower == "Tower G" && TowerG == 0)
            {
                 TowerG++;
                 obj['project_name'] = "Greenfield Phase 2";
                 obj['tower_name'] = "Tower G";
                 obj['total_unit'] = TowerG;
                 arr.push(obj);                  
            }
            else if(otherActiviytyValue[i].tower == "Tower H" && TowerH == 0)
            {
              TowerH++;
              obj['tower_name'] = "Tower H";
              obj['total_unit'] = TowerH;
              arr.push(obj);                  
            }
            else if(otherActiviytyValue[i].tower == "Tower J" && TowerJ == 0)
            {
              TowerJ++;
              obj['tower_name'] = "Tower J";
              obj['total_unit'] = TowerJ;
              arr.push(obj);                  
            }
        }          
        else if(otherActiviytyValue[i].project_name == "Southern Crest")
        {
          var obj = {};
          if(otherActiviytyValue[i].tower == "Tower A" && TowerSA == 0)
          {
            TowerSA++;
            obj['project_name'] = "Southern Crest";
            obj['tower_name'] = "Tower A";
            obj['total_unit'] = TowerSA;
            arr.push(obj);
          }
          else if(otherActiviytyValue[i].tower == "Tower B" && TowerSB == 0)
          {
              TowerSB++;
              obj['tower_name'] = "Tower B";
              obj['total_unit'] = TowerSB;
              arr.push(obj);                  
          }
          else if(otherActiviytyValue[i].tower == "Tower C" && TowerSC == 0)
          {
              TowerSC++;
              obj['tower_name'] = "Tower C";
              obj['total_unit'] = TowerSC;
              arr.push(obj);                  
          }
        }
        else if(otherActiviytyValue[i].project_name == "Luxor")
        {
          var obj = {};
          if(otherActiviytyValue[i].tower == "Tower A" && TowerLA == 0)
          {
            TowerLA++;
            obj['project_name'] = "Luxor";
            obj['tower_name'] = "Tower A";
            obj['total_unit'] = TowerLA;
            arr.push(obj);
          }
          else if(otherActiviytyValue[i].tower == "Tower B" && TowerLB == 0)
          {
            TowerLB++;
            obj['tower_name'] = "Tower B";
            obj['total_unit'] = TowerLB;
            arr.push(obj);                  
          }
          else if(otherActiviytyValue[i].tower == "Tower C" && TowerLC == 0)
          {
            TowerLC++;
            obj['tower_name'] = "Tower C";
            obj['total_unit'] = TowerLC;
            arr.push(obj);                  
          }
          else if(otherActiviytyValue[i].tower == "Tower D" && TowerLD == 0)
          {
            TowerLD++;
            obj['tower_name'] = "Tower D";
            obj['total_unit'] = TowerLD;
            arr.push(obj);                  
          }
          else if(otherActiviytyValue[i].tower == "Tower E" && TowerLE == 0)
          {
            TowerLE++;
            obj['tower_name'] = "Tower E";
            obj['total_unit'] = TowerLE;
            arr.push(obj);                  
          }
          else if(otherActiviytyValue[i].tower == "Tower F" && TowerLF == 0)
          {
            TowerLF++;
            obj['tower_name'] = "Tower F";
            obj['total_unit'] = TowerLF;
            arr.push(obj);                  
          }
          else if(otherActiviytyValue[i].tower == "Tower G" && TowerLG == 0)
          {
            TowerLG++;
            obj['tower_name'] = "Tower G";
            obj['total_unit'] = TowerLG;
            arr.push(obj);                  
          }
          else if(otherActiviytyValue[i].tower == "Tower H" && TowerLH == 0)
          {
            TowerLH++;
            obj['tower_name'] = "Tower H";
            obj['total_unit'] = TowerLH;
            arr.push(obj);                  
          }
        }
      }
      
        //loop to checkmilestone status
       for(var j=1;j<otherActiviytyValue.length;j++)
        {
                 // milestone activities
                 var comp = {};
                 var uncomp = {};
     
     
                 if(otherActiviytyValue[j].project_name == "Panorama Hills" && otherActiviytyValue[j].tower == "2A")
                 {   
                   if(otherActiviytyValue[j].activity_status == "1")
                   {
                     var completed = 1;
                     comp['milestone'] = otherActiviytyValue[j].activity_name;
                     comp['count'] = completed;
                     Tower2A.push(comp);            
                   }
                   else if(otherActiviytyValue[j].activity_status == "0")
                   {
                     var uncompleted = 0;
                     uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                     uncomp['count'] = uncompleted;
                     Tower2A.push(uncomp);        
                   }
                 }    
                 else if(otherActiviytyValue[j].project_name == "Panorama Hills" && otherActiviytyValue[j].tower == "2B")
                 {   
                   if(otherActiviytyValue[j].activity_status == "1")
                   {
                     var completed = 1;
                     comp['milestone'] = otherActiviytyValue[j].activity_name;
                     comp['count'] = completed;
                     Tower2B.push(comp);            
                   }
                   else if(otherActiviytyValue[j].activity_status == "0")
                   {
                     var uncompleted = 0;
                     uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                     uncomp['count'] = uncompleted;
                     Tower2B.push(uncomp);        
                   }
                 }
                 else if(otherActiviytyValue[j].project_name == "Panorama Hills" && otherActiviytyValue[j].tower == "2C")
                 {   
                   if(otherActiviytyValue[j].activity_status == "1")
                   {
                     var completed = 1;
                     comp['milestone'] = otherActiviytyValue[j].activity_name;
                     comp['count'] = completed;
                     Tower2C.push(comp);            
                   }
                   else if(otherActiviytyValue[j].activity_status == "0")
                   {
                     var uncompleted = 0;
                     uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                     uncomp['count'] = uncompleted;
                     Tower2C.push(uncomp);        
                   }
                 }    
                 else if(otherActiviytyValue[j].project_name == "Panorama Hills" && otherActiviytyValue[j].tower == "2D")
                 {   
                   if(otherActiviytyValue[j].activity_status == "1")
                   {
                     var completed = 1;
                     comp['milestone'] = otherActiviytyValue[j].activity_name;
                     comp['count'] = completed;
                     Tower2D.push(comp);            
                   }
                   else if(otherActiviytyValue[j].activity_status == "0")
                   {
                     var uncompleted = 0;
                     uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                     uncomp['count'] = uncompleted;
                     Tower2D.push(uncomp);        
                   }
                 }
                 else if(otherActiviytyValue[j].project_name == "Greenfield Phase 1" && otherActiviytyValue[j].tower == "Tower A")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr6.push(comp);
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr6.push(uncomp);        
                     }  
                 }
                 else if(otherActiviytyValue[j].project_name == "Greenfield Phase 1" && otherActiviytyValue[j].tower == "Tower B")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr7.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr7.push(uncomp);        
                     }  
                 }
                 else if(otherActiviytyValue[j].project_name == "Greenfield Phase 1" && otherActiviytyValue[j].tower == "Tower C")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr8.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr8.push(uncomp);        
                     }  
                 }
                 else if(otherActiviytyValue[j].project_name == "Greenfield Phase 1" && otherActiviytyValue[j].tower == "Tower D")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr9.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr9.push(uncomp);        
                     }  
                 }
                 else if(otherActiviytyValue[j].project_name == "Greenfield Phase 1" && otherActiviytyValue[j].tower == "Tower E")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr10.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr10.push(uncomp);        
                     }  
                 }
                 else if(otherActiviytyValue[j].project_name == "Greenfield Phase 1" && otherActiviytyValue[j].tower == "Tower F")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr11.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr11.push(uncomp);        
                     }  
                 }
                 else if(otherActiviytyValue[j].project_name == "Greenfield Phase 2" && otherActiviytyValue[j].tower == "Tower G")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr23.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr23.push(uncomp);        
                     }   
                 }
                 else if(otherActiviytyValue[j].project_name == "Greenfield Phase 2" && otherActiviytyValue[j].tower == "Tower H")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr24.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr24.push(uncomp);        
                     }   
                 }
                 else if(otherActiviytyValue[j].project_name == "Greenfield Phase 2" && otherActiviytyValue[j].tower == "Tower J")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr25.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr25.push(uncomp);        
                     }   
                 }
                 else if(otherActiviytyValue[j].project_name == "Southern Crest" && otherActiviytyValue[j].tower == "Tower A")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr12.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr12.push(uncomp);        
                     }  
                 }
                 else if(otherActiviytyValue[j].project_name == "Southern Crest" && otherActiviytyValue[j].tower == "Tower B")
                 {
                   if(otherActiviytyValue[j].activity_status == "1")
                   {
                     var completed = 1;
                     comp['milestone'] = otherActiviytyValue[j].activity_name;
                     comp['count'] = completed;
                     arr13.push(comp);            
                   }
                   else if(otherActiviytyValue[j].activity_status == "0")
                   {
                     var uncompleted = 0;
                     uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                     uncomp['count'] = uncompleted;
                     arr13.push(uncomp);        
                   }              
                 }
                 else if(otherActiviytyValue[j].project_name == "Southern Crest" && otherActiviytyValue[j].tower == "Tower C")
                 {
                   if(otherActiviytyValue[j].activity_status == "1")
                   {
                     var completed = 1;
                     comp['milestone'] = otherActiviytyValue[j].activity_name;
                     comp['count'] = completed;
                     arr14.push(comp);            
                   }
                   else if(otherActiviytyValue[j].activity_status == "0")
                   {
                     var uncompleted = 0;
                     uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                     uncomp['count'] = uncompleted;
                     arr14.push(uncomp);        
                   }
                 }
                 else if(otherActiviytyValue[j].project_name == "Luxor" && otherActiviytyValue[j].tower == "Tower A")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr15.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr15.push(uncomp);        
                     }   
                 }
                 else if(otherActiviytyValue[j].project_name == "Luxor" && otherActiviytyValue[j].tower == "Tower B")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr16.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr16.push(uncomp);        
                     }   
                 }    
                 else if(otherActiviytyValue[j].project_name == "Luxor" && otherActiviytyValue[j].tower == "Tower C")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr17.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr17.push(uncomp);        
                     }   
                 }
                 else if(otherActiviytyValue[j].project_name == "Luxor" && otherActiviytyValue[j].tower == "Tower D")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = structureActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr18.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr18.push(uncomp);        
                     }   
                 }    
                 else if(otherActiviytyValue[j].project_name == "Luxor" && otherActiviytyValue[j].tower == "Tower E")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr19.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr19.push(uncomp);        
                     }   
                 }
                 else if(otherActiviytyValue[j].project_name == "Luxor" && otherActiviytyValue[j].tower == "Tower F")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr20.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr20.push(uncomp);        
                     }   
                 }
                 else if(otherActiviytyValue[j].project_name == "Luxor" && otherActiviytyValue[j].tower == "Tower G")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr21.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr21.push(uncomp);        
                     }   
                 }
                 else if(otherActiviytyValue[j].project_name == "Luxor" && otherActiviytyValue[j].tower == "Tower H")
                 {
                     if(otherActiviytyValue[j].activity_status == "1")
                     {
                       var completed = 1;
                       comp['milestone'] = otherActiviytyValue[j].activity_name;
                       comp['count'] = completed;
                       arr22.push(comp);            
                     }
                     else if(otherActiviytyValue[j].activity_status == "0")
                     {
                       var uncompleted = 0;
                       uncomp['milestone'] = otherActiviytyValue[j].activity_name;
                       uncomp['count'] = uncompleted;
                       arr22.push(uncomp);        
                     }   
                 }     
       }

    //other Activiyty excel for panorama hills
     generateExcelOtherPanorama(arr,Tower2A,Tower2B,Tower2C,Tower2D); 

     //other Activiyty excel for greenfield phase 1
     generateExcelOtherPhase1(arr,arr6,arr7,arr8,arr9,arr10,arr11);

     //other Activiyty excel for greenfield phase 2
     generateExcelOtherPhase2(arr,arr23,arr24,arr25);

     //other Activiyty excel for southren crest
     generateExcelOtherSC(arr,arr12,arr13,arr14);

     //other activity for Luxor
     generateExcelOtherLuxor(arr,arr15,arr16,arr17,arr18,arr19,arr20,arr21,arr22);

     res.status(200).send({
        error :false,
        Array :arr,
    });

  });
});


//api for structure activity for firebase
router.post('/structure_activities',(req, res, next) => {
    const dbRef = firebase.database().ref("structure_activities");

    var arr  = new Array();//0-3 vizag, 4-9 greenfield phase 1, 10-12 southren crest, 13-20 Luxor
    var arr2 = new Array();
    var arr3 = new Array();
    var arr4 = new Array();
    var arr5 = new Array();
    var garr = new Array();

    //arr for greenphase-1 
    var arr6 = new Array();
    var arr7 = new Array();
    var arr8 = new Array();
    var arr9 = new Array();
    var arr10 = new Array();
    var arr11 = new Array();

    //arr for Southren crest 
    var arr12 = new Array();
    var arr13 = new Array();
    var arr14 = new Array();

    //arr for Luxor
    var arr15 = new Array();
    var arr16 = new Array();
    var arr17 = new Array();
    var arr18 = new Array();
    var arr19 = new Array();
    var arr20 = new Array();
    var arr21 = new Array();
    var arr22 = new Array();

    //arr for greenfield phase2
    var arr23 =  new Array();
    var arr24 =  new Array();
    var arr25 =  new Array();

    var obj1 = {};
    var obj2 = {};
    var obj3 = {};
    var obj4 = {};



    dbRef.once("value").then(snap => {
        //panorama hills
        var T2ACount = 0;
        var T2BCount = 0;
        var T2CCount = 0;
        var T2DCount = 0;
        var T2ECount = 0;
        var T2FCount = 0;

        //greenfield pahse1
        var TowerA = 0;
        var TowerB = 0;
        var TowerC = 0;
        var TowerD = 0;
        var TowerE = 0;
        var TowerF = 0;

        //greenfield pahse2
        var TowerG = 0;
        var TowerH = 0;
        var TowerJ = 0;

        //southrencrest variable
        var TowerSA = 0;
        var TowerSB = 0;
        var TowerSC = 0;

        //LUXOR Varirables
        var TowerLA = 0;
        var TowerLB = 0;
        var TowerLC = 0;
        var TowerLD = 0;
        var TowerLE = 0;
        var TowerLF = 0;
        var TowerLG = 0;
        var TowerLH = 0;
        
        var structureActiviytyValue = snap.val();


        //loop for getting total number of tower in the flat
        for(var i=1;i<structureActiviytyValue.length;i++)
        {
          if(structureActiviytyValue[i].project_name == "Panorama Hills")
          {   
            if(structureActiviytyValue[i].tower == "2A" && T2ACount == 0)
            {
                T2ACount++;
                obj1['tower_name'] = "Tower 2A";
                obj1['total_unit'] = T2ACount;
                arr.push(obj1);
            }
            else if(structureActiviytyValue[i].tower == "2B" && T2BCount == 0)
            {
                T2BCount++;
                obj2['tower_name'] = "Tower 2B";
                obj2['total_unit'] = T2BCount;
                arr.push(obj2);   
            }
            else if(structureActiviytyValue[i].tower == "2C" && T2CCount == 0)
            {
                T2CCount++;
                obj3['tower_name'] = "Tower 2C";
                obj3['total_unit'] = T2CCount;
                arr.push(obj3);  
            }
            else if(structureActiviytyValue[i].tower == "2D" && T2DCount == 0)
            {
                T2DCount++;
                obj4['tower_name'] = "Tower 2D";
                obj4['total_unit'] = T2DCount; 
                arr.push(obj4);  
            }
          }
          else if(structureActiviytyValue[i].project_name == "Greenfield Phase 1")
          {
              var obj = {};

            if(structureActiviytyValue[i].tower == "Tower A" && TowerA == 0)
            {
              TowerA++;
              obj['tower_name'] = "Tower A";
              obj['total_unit'] = TowerA;
              arr.push(obj);
            }
            else if(structureActiviytyValue[i].tower == "Tower B" && TowerB == 0)
            {
                TowerB++;
                obj['tower_name'] = "Tower B";
                obj['total_unit'] = TowerB;
                arr.push(obj);                  
            }
            else if(structureActiviytyValue[i].tower == "Tower C" && TowerC == 0)
            {
                TowerC++;
                obj['tower_name'] = "Tower C";
                obj['total_unit'] = TowerC;
                arr.push(obj);                                  
            }
            else if(structureActiviytyValue[i].tower == "Tower D" && TowerD == 0)
            {
                TowerD++;
                obj['tower_name'] = "Tower D";
                obj['total_unit'] = TowerD;
                arr.push(obj);                                  
            }
            else if(structureActiviytyValue[i].tower == "Tower E" && TowerE == 0)
            {
                TowerE++;
                obj['tower_name'] = "Tower E";
                obj['total_unit'] = TowerE;
                arr.push(obj);                                  
            }
            else if(structureActiviytyValue[i].tower == "Tower F" && TowerF == 0)
            {
                TowerF++;
                obj['tower_name'] = "Tower F";
                obj['total_unit'] = TowerF;
                arr.push(obj);                                  
            }
          }
          else if(structureActiviytyValue[i].project_name == "Southern Crest")
          {
            var obj = {};
            if(structureActiviytyValue[i].tower == "Tower A" && TowerSA == 0)
            {
              TowerSA++;
              obj['tower_name'] = "Tower A";
              obj['total_unit'] = TowerSA;
              arr.push(obj);
            }
            else if(structureActiviytyValue[i].tower == "Tower B" && TowerSB == 0)
            {
                TowerSB++;
                obj['tower_name'] = "Tower B";
                obj['total_unit'] = TowerSB;
                arr.push(obj);                  
            }
            else if(structureActiviytyValue[i].tower == "Tower C" && TowerSC == 0)
            {
                TowerSC++;
                obj['tower_name'] = "Tower C";
                obj['total_unit'] = TowerSC;
                arr.push(obj);                  
            }
          }
          else if(structureActiviytyValue[i].project_name == "Luxor")
          {
            var obj = {};
            if(structureActiviytyValue[i].tower == "Tower A" && TowerLA == 0)
            {
              TowerLA++;
              obj['tower_name'] = "Tower A";
              obj['total_unit'] = TowerLA;
              arr.push(obj);
            }
            else if(structureActiviytyValue[i].tower == "Tower B" && TowerLB == 0)
            {
              TowerLB++;
              obj['tower_name'] = "Tower B";
              obj['total_unit'] = TowerLB;
              arr.push(obj);                  
            }
            else if(structureActiviytyValue[i].tower == "Tower C" && TowerLC == 0)
            {
              TowerLC++;
              obj['tower_name'] = "Tower C";
              obj['total_unit'] = TowerLC;
              arr.push(obj);                  
            }
            else if(structureActiviytyValue[i].tower == "Tower D" && TowerLD == 0)
            {
              TowerLD++;
              obj['tower_name'] = "Tower D";
              obj['total_unit'] = TowerLD;
              arr.push(obj);                  
            }
            else if(structureActiviytyValue[i].tower == "Tower E" && TowerLE == 0)
            {
              TowerLE++;
              obj['tower_name'] = "Tower E";
              obj['total_unit'] = TowerLE;
              arr.push(obj);                  
            }
            else if(structureActiviytyValue[i].tower == "Tower F" && TowerLF == 0)
            {
              TowerLF++;
              obj['tower_name'] = "Tower F";
              obj['total_unit'] = TowerLF;
              arr.push(obj);                  
            }
            else if(structureActiviytyValue[i].tower == "Tower G" && TowerLG == 0)
            {
              TowerLG++;
              obj['tower_name'] = "Tower G";
              obj['total_unit'] = TowerLG;
              arr.push(obj);                  
            }
            else if(structureActiviytyValue[i].tower == "Tower H" && TowerLH == 0)
            {
              TowerLH++;
              obj['tower_name'] = "Tower H";
              obj['total_unit'] = TowerLH;
              arr.push(obj);                  
            }
          }
          else if(structureActiviytyValue[i].project_name == "Greenfield Phase 2")
          {

              var obj = {};
              if(structureActiviytyValue[i].tower == "Tower G" && TowerG == 0)
              {
                   TowerG++;
                   obj['tower_name'] = "Tower G";
                   obj['total_unit'] = TowerG;
                   garr.push(obj);                  
              }
              else if(structureActiviytyValue[i].tower == "Tower H" && TowerH == 0)
              {
                TowerH++;
                obj['tower_name'] = "Tower H";
                obj['total_unit'] = TowerH;
                garr.push(obj);                  
              }
              else if(structureActiviytyValue[i].tower == "Tower J" && TowerJ == 0)
              {
                TowerJ++;
                obj['tower_name'] = "Tower J";
                obj['total_unit'] = TowerJ;
                garr.push(obj);                  
              }
          }
        }


        //loop to checkmilestone status
        for(var j=1;j<structureActiviytyValue.length;j++)
        {
            // milestone activities
            var comp = {};
            var uncomp = {};


            if(structureActiviytyValue[j].project_name == "Panorama Hills" && structureActiviytyValue[j].tower == "2A")
            {   
              if(structureActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = structureActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr2.push(comp);            
              }
              else if(structureActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr2.push(uncomp);        
              }
            }    
            else if(structureActiviytyValue[j].project_name == "Panorama Hills" && structureActiviytyValue[j].tower == "2B")
            {   
              if(structureActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = structureActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr3.push(comp);            
              }
              else if(structureActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr3.push(uncomp);        
              }
            }
            else if(structureActiviytyValue[j].project_name == "Panorama Hills" && structureActiviytyValue[j].tower == "2C")
            {   
              if(structureActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = structureActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr4.push(comp);            
              }
              else if(structureActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr4.push(uncomp);        
              }
            }    
            else if(structureActiviytyValue[j].project_name == "Panorama Hills" && structureActiviytyValue[j].tower == "2D")
            {   
              if(structureActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = structureActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr5.push(comp);            
              }
              else if(structureActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr5.push(uncomp);        
              }
            }
            else if(structureActiviytyValue[j].project_name == "Greenfield Phase 1" && structureActiviytyValue[j].tower == "Tower A")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr6.push(comp);
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr6.push(uncomp);        
                }  
            }
            else if(structureActiviytyValue[j].project_name == "Greenfield Phase 1" && structureActiviytyValue[j].tower == "Tower B")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr7.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr7.push(uncomp);        
                }  
            }
            else if(structureActiviytyValue[j].project_name == "Greenfield Phase 1" && structureActiviytyValue[j].tower == "Tower C")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr8.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr8.push(uncomp);        
                }  
            }
            else if(structureActiviytyValue[j].project_name == "Greenfield Phase 1" && structureActiviytyValue[j].tower == "Tower D")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr9.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr9.push(uncomp);        
                }  
            }
            else if(structureActiviytyValue[j].project_name == "Greenfield Phase 1" && structureActiviytyValue[j].tower == "Tower E")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr10.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr10.push(uncomp);        
                }  
            }
            else if(structureActiviytyValue[j].project_name == "Greenfield Phase 1" && structureActiviytyValue[j].tower == "Tower F")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr11.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr11.push(uncomp);        
                }  
            }
            else if(structureActiviytyValue[j].project_name == "Greenfield Phase 2" && structureActiviytyValue[j].tower == "Tower G")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr23.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr23.push(uncomp);        
                }   
            }
            else if(structureActiviytyValue[j].project_name == "Greenfield Phase 2" && structureActiviytyValue[j].tower == "Tower H")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr24.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr24.push(uncomp);        
                }   
            }
            else if(structureActiviytyValue[j].project_name == "Greenfield Phase 2" && structureActiviytyValue[j].tower == "Tower J")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr25.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr25.push(uncomp);        
                }   
            }
            else if(structureActiviytyValue[j].project_name == "Southern Crest" && structureActiviytyValue[j].tower == "Tower A")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr12.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr12.push(uncomp);        
                }  
            }
            else if(structureActiviytyValue[j].project_name == "Southern Crest" && structureActiviytyValue[j].tower == "Tower B")
            {
              if(structureActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = structureActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr13.push(comp);            
              }
              else if(structureActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr13.push(uncomp);        
              }              
            }
            else if(structureActiviytyValue[j].project_name == "Southern Crest" && structureActiviytyValue[j].tower == "Tower C")
            {
              if(structureActiviytyValue[j].activity_status == "1")
              {
                var completed = 1;
                comp['milestone'] = structureActiviytyValue[j].activity_name;
                comp['count'] = completed;
                arr14.push(comp);            
              }
              else if(structureActiviytyValue[j].activity_status == "0")
              {
                var uncompleted = 0;
                uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                uncomp['count'] = uncompleted;
                arr14.push(uncomp);        
              }
            }
            else if(structureActiviytyValue[j].project_name == "Luxor" && structureActiviytyValue[j].tower == "Tower A")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr15.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr15.push(uncomp);        
                }   
            }
            else if(structureActiviytyValue[j].project_name == "Luxor" && structureActiviytyValue[j].tower == "Tower B")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr16.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr16.push(uncomp);        
                }   
            }    
            else if(structureActiviytyValue[j].project_name == "Luxor" && structureActiviytyValue[j].tower == "Tower C")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr17.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr17.push(uncomp);        
                }   
            }
            else if(structureActiviytyValue[j].project_name == "Luxor" && structureActiviytyValue[j].tower == "Tower D")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr18.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr18.push(uncomp);        
                }   
            }    
            else if(structureActiviytyValue[j].project_name == "Luxor" && structureActiviytyValue[j].tower == "Tower E")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr19.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr19.push(uncomp);        
                }   
            }
            else if(structureActiviytyValue[j].project_name == "Luxor" && structureActiviytyValue[j].tower == "Tower F")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr20.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr20.push(uncomp);        
                }   
            }
            else if(structureActiviytyValue[j].project_name == "Luxor" && structureActiviytyValue[j].tower == "Tower G")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr21.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr21.push(uncomp);        
                }   
            }
            else if(structureActiviytyValue[j].project_name == "Luxor" && structureActiviytyValue[j].tower == "Tower H")
            {
                if(structureActiviytyValue[j].activity_status == "1")
                {
                  var completed = 1;
                  comp['milestone'] = structureActiviytyValue[j].activity_name;
                  comp['count'] = completed;
                  arr22.push(comp);            
                }
                else if(structureActiviytyValue[j].activity_status == "0")
                {
                  var uncompleted = 0;
                  uncomp['milestone'] = structureActiviytyValue[j].activity_name;
                  uncomp['count'] = uncompleted;
                  arr22.push(uncomp);        
                }   
            }
            
        }
 
       /*
        *Excel function for panorama hills(vizag)
        */
        generateExcel(arr,arr2,arr3,arr4,arr5);
       
        
       /*
        *Excel function for greenfield phase1
        */
        generateExcelPhase1(arr,arr6,arr7,arr8,arr9,arr10,arr11);

       /*
        *Excel function for southren crest
        */
        generateExcelSC(arr,arr12,arr13,arr14);


        /*
         *Excel function for Luxor 
         */
        generateExcelLuxor(arr,arr15,arr16,arr17,arr18,arr19,arr20,arr21,arr22);
        
        /*
         *Excel function for Luxor 
         */
        generateExcelPhase2(garr,arr23,arr24,arr25);

        res.status(200).send({
            error :false,
            Array :arr,
            Array1:garr,
        });

    }).catch(error => {
        console.log("error".error)
    });
});

//api for commonlobby activities
router.post('/commonlobbyActivities',(req,res, next) =>{

    const dbRef = firebase.database().ref("commonlobby_activities");

    //array to keep count of number of floors
    var floor_count = new Array();

    //to keep track of count

    //panorama hills
    var Tower2A = 0;
    var Tower2B = 0;
    var Tower2C = 0;    
    var Tower2D = 0;    
    
    //greenfield phase1
    var TowerA = 0;
    var TowerB = 0;
    var TowerC = 0;
    var TowerD = 0;
    var TowerE = 0;
    var TowerF = 0;
          
    //greenfield phase2
    var TowerG = 0;
    var TowerH = 0;
    var TowerJ = 0;
          
    //southrencrest variable
    var TowerSA = 0;
    var TowerSB = 0;
    var TowerSC = 0;
          
    //LUXOR Varirables
    var TowerLA = 0;
    var TowerLB = 0;
    var TowerLC = 0;
    var TowerLD = 0;
    var TowerLE = 0;
    var TowerLF = 0;
    var TowerLG = 0;
    var TowerLH = 0;
    

    //array for panorama hills
    var arr1 = new Array();
    var arr2 = new Array();
    var arr3 = new Array();
    var arr4 = new Array();

    //arr for greenphase-1 
    var arr6 = new Array();//greenphase-1  Tower A
    var arr7 = new Array();//greenphase-1  Tower B
    var arr8 = new Array();//greenphase-1  Tower C
    var arr9 = new Array();//greenphase-1  Tower D
    var arr10 = new Array();//greenphase-1  Tower E
    var arr11 = new Array();//greenphase-1  Tower F
    
    //arr for greenphase-2 
    var arr12 = new Array();
    var arr13 = new Array();
    var arr14 = new Array();
    
    //arr for Southrencrest
    var arr15 = new Array();
    var arr16 = new Array();
    var arr17 = new Array();

    //arr for luxor
    var arr18 = new Array();
    var arr19 = new Array();
    var arr20 = new Array();
    var arr21 = new Array();
    var arr22 = new Array();
    var arr23 =  new Array();
    var arr24 =  new Array();
    var arr25 =  new Array();
    

    dbRef.once("value").then(snap =>{
        var commonlobbyActivitiesValue = snap.val();
        var length = commonlobbyActivitiesValue.length;
        for(var i=1;i<commonlobbyActivitiesValue.length;i++)
        {
            var obj = {};

            if(commonlobbyActivitiesValue[i].project_name == "Panorama Hills")
            {
               if(commonlobbyActivitiesValue[i].tower == "2A")
               {
                    Tower2A++;
                    var tower  = commonlobbyActivitiesValue[i].tower;
                    if(floor_count.length != 0)
                    {
                      var result = checkObject(tower,floor_count,Tower2A);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = Tower2A;
                      }
                    }
                    else
                    {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = Tower2A;
                       floor_count.push(obj);
                    }
               }
               else if(commonlobbyActivitiesValue[i].tower == "2B")
               {
                     Tower2B++;
                     var tower  = commonlobbyActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,Tower2B);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = Tower2B;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = Tower2B;
                      floor_count.push(obj);
                     }
                }
                else if(commonlobbyActivitiesValue[i].tower == "2C")
                {
                    Tower2C++;
                    var tower  = commonlobbyActivitiesValue[i].tower;
                    var result = checkObject(tower,floor_count,Tower2C);
                    if(result != "false")
                    {
                      floor_count[result]['total_unit'] = Tower2C;
                    }
                    else
                    {
                     obj['tower_name'] = tower,
                     obj['total_unit'] = Tower2C;
                     floor_count.push(obj);
                    }
               }
               else if(commonlobbyActivitiesValue[i].tower == "2D")
               {
                   Tower2D++;
                   var tower  = commonlobbyActivitiesValue[i].tower;
                   var result = checkObject(tower,floor_count,Tower2D);
                   if(result != "false")
                   {
                     floor_count[result]['total_unit'] = Tower2D;
                   }
                   else
                   {
                    obj['tower_name'] = tower,
                    obj['total_unit'] = Tower2D;
                    floor_count.push(obj);
                   }
              }
            }
            else if(commonlobbyActivitiesValue[i].project_name == "Greenfield Phase 1")
            {
                if(commonlobbyActivitiesValue[i].tower == "Tower A")
                {
                     TowerA++;
                     var tower  = commonlobbyActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerA);
                     if(result != "false")
                     {
                         floor_count[result]['total_unit'] = TowerA;
                     }
                     else
                     {
                        obj['Project Name'] = "Greenfield Phase 1",
                        obj['tower_name'] = tower,
                        obj['total_unit'] = TowerA;
                        floor_count.push(obj);
                     }
                }
                else if(commonlobbyActivitiesValue[i].tower == "Tower B")
                {
                      TowerB++;
                      var tower  = commonlobbyActivitiesValue[i].tower;
                      var result = checkObject(tower,floor_count,TowerB);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = TowerB;
                      }
                      else
                      {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = TowerB;
                       floor_count.push(obj);
                      }
                 }
                 else if(commonlobbyActivitiesValue[i].tower == "Tower C")
                 {
                     TowerC++;
                     var tower  = commonlobbyActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerC);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = TowerC;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = TowerC;
                      floor_count.push(obj);
                     }
                }
                else if(commonlobbyActivitiesValue[i].tower == "Tower D")
                {
                    TowerD++;
                    var tower  = commonlobbyActivitiesValue[i].tower;
                    var result = checkObject(tower,floor_count,TowerD);
                    if(result != "false")
                    {
                      floor_count[result]['total_unit'] = TowerD;
                    }
                    else
                    {
                     obj['tower_name'] = tower,
                     obj['total_unit'] = TowerD;
                     floor_count.push(obj);
                    }
               }
               else if(commonlobbyActivitiesValue[i].tower == "Tower E")
               {
                   TowerE++;
                   var tower  = commonlobbyActivitiesValue[i].tower;
                   var result = checkObject(tower,floor_count,TowerE);
                   if(result != "false")
                   {
                     floor_count[result]['total_unit'] = TowerE;
                   }
                   else
                   {
                    obj['tower_name'] = tower,
                    obj['total_unit'] = TowerE;
                    floor_count.push(obj);
                   }
               }
               else if(commonlobbyActivitiesValue[i].tower == "Tower F")
               {
                   TowerF++;
                   var tower  = commonlobbyActivitiesValue[i].tower;
                   var result = checkObject(tower,floor_count,TowerE);
                   if(result != "false")
                   {
                     floor_count[result]['total_unit'] = TowerF;
                   }
                   else
                   {
                    obj['tower_name'] = tower,
                    obj['total_unit'] = TowerF;
                    floor_count.push(obj);
                   }
               }
            }
            else if(commonlobbyActivitiesValue[i].project_name == "Greenfield Phase 2")
            {
                if(commonlobbyActivitiesValue[i].tower == "Tower G")
                {
                     TowerG++;
                     var tower  = commonlobbyActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerG);
                     if(result != "false")
                     {
                         floor_count[result]['total_unit'] = TowerG;
                     }
                     else
                     {
                        obj['Project Name'] = "Greenfield Phase 2",
                        obj['tower_name'] = tower,
                        obj['total_unit'] = TowerG;
                        floor_count.push(obj);
                     }
                }
                else if(commonlobbyActivitiesValue[i].tower == "Tower H")
                {
                      TowerH++;
                      var tower  = commonlobbyActivitiesValue[i].tower;
                      var result = checkObject(tower,floor_count,TowerH);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = TowerH;
                      }
                      else
                      {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = TowerH;
                       floor_count.push(obj);
                      }
                 }
                 else if(commonlobbyActivitiesValue[i].tower == "Tower J")
                 {
                     TowerJ++;
                     var tower  = commonlobbyActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerJ);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = TowerJ;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = TowerJ;
                      floor_count.push(obj);
                     }
                }
            }
            else if(commonlobbyActivitiesValue[i].project_name == "Southern Crest")
            {
                if(commonlobbyActivitiesValue[i].tower == "Tower A")
                {
                     TowerSA++;
                     var tower  = commonlobbyActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerSA);
                     if(result != "false")
                     {
                         floor_count[result]['total_unit'] = TowerSA;
                     }
                     else
                     {
                        obj['Project Name'] = "Southren Crest",
                        obj['tower_name'] = tower,
                        obj['total_unit'] = TowerSA;
                        floor_count.push(obj);
                     }
                }
                else if(commonlobbyActivitiesValue[i].tower == "Tower B")
                {
                      TowerSB++;
                      var tower  = commonlobbyActivitiesValue[i].tower;
                      var result = checkObject(tower,floor_count,TowerSB);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = TowerSB;
                      }
                      else
                      {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = TowerSB;
                       floor_count.push(obj);
                      }
                 }
                 else if(commonlobbyActivitiesValue[i].tower == "Tower C")
                 {
                     TowerSC++;
                     var tower  = commonlobbyActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerSC);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = TowerSC;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = TowerSC;
                      floor_count.push(obj);
                     }
                }
            }
            else if(commonlobbyActivitiesValue[i].project_name == "Luxor")
            {
                if(commonlobbyActivitiesValue[i].tower == "Tower A")
                {
                     TowerLA++;
                     var tower  = commonlobbyActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerLA);
                     if(result != "false")
                     {
                         floor_count[result]['total_unit'] = TowerLA;
                     }
                     else
                     {
                        obj['Project Name'] = "Luxor",
                        obj['tower_name'] = tower,
                        obj['total_unit'] = TowerLA;
                        floor_count.push(obj);
                     }
                }
                else if(commonlobbyActivitiesValue[i].tower == "Tower B")
                {
                      TowerLB++;
                      var tower  = commonlobbyActivitiesValue[i].tower;
                      var result = checkObject(tower,floor_count,TowerLB);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = TowerLB;
                      }
                      else
                      {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = TowerLB;
                       floor_count.push(obj);
                      }
                 }
                 else if(commonlobbyActivitiesValue[i].tower == "Tower C")
                 {
                     TowerLC++;
                     var tower  = commonlobbyActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerLC);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = TowerLC;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = TowerLC;
                      floor_count.push(obj);
                     }
                }
                else if(commonlobbyActivitiesValue[i].tower == "Tower D")
                {
                    TowerLD++;
                    var tower  = commonlobbyActivitiesValue[i].tower;
                    var result = checkObject(tower,floor_count,TowerLD);
                    if(result != "false")
                    {
                      floor_count[result]['total_unit'] = TowerLD;
                    }
                    else
                    {
                     obj['tower_name'] = tower,
                     obj['total_unit'] = TowerLD;
                     floor_count.push(obj);
                    }
               }
               else if(commonlobbyActivitiesValue[i].tower == "Tower E")
               {
                   TowerLE++;
                   var tower  = commonlobbyActivitiesValue[i].tower;
                   var result = checkObject(tower,floor_count,TowerLE);
                   if(result != "false")
                   {
                     floor_count[result]['total_unit'] = TowerLE;
                   }
                   else
                   {
                    obj['tower_name'] = tower,
                    obj['total_unit'] = TowerLE;
                    floor_count.push(obj);
                   }
              }
              else if(commonlobbyActivitiesValue[i].tower == "Tower F")
              {
                  TowerLF++;
                  var tower  = commonlobbyActivitiesValue[i].tower;
                  var result = checkObject(tower,floor_count,TowerLF);
                  if(result != "false")
                  {
                    floor_count[result]['total_unit'] = TowerLF;
                  }
                  else
                  {
                   obj['tower_name'] = tower,
                   obj['total_unit'] = TowerLF;
                   floor_count.push(obj);
                  }
             }
             else if(commonlobbyActivitiesValue[i].tower == "Tower G")
             {
                 TowerLG++;
                 var tower  = commonlobbyActivitiesValue[i].tower;
                 var result = checkObject(tower,floor_count,TowerLG);
                 if(result != "false")
                 {
                   floor_count[result]['total_unit'] = TowerLG;
                 }
                 else
                 {
                  obj['tower_name'] = tower,
                  obj['total_unit'] = TowerLG;
                  floor_count.push(obj);
                 }
             }
             else if(commonlobbyActivitiesValue[i].tower == "Tower H")
             {
                 TowerLH++;
                 var tower  = commonlobbyActivitiesValue[i].tower;
                 var result = checkObject(tower,floor_count,TowerLH);
                 if(result != "false")
                 {
                   floor_count[result]['total_unit'] = TowerLH;
                 }
                 else
                 {
                  obj['tower_name'] = tower,
                  obj['total_unit'] = TowerLH;
                  floor_count.push(obj);
                 }
             }
            }
            
            //to access activities
            if(commonlobbyActivitiesValue[i].project_name =="Panorama Hills" && commonlobbyActivitiesValue[i].tower == "2A")
            {  
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr1.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr1);
                           
                           if(result != "false")
                           {
                               var data = arr1[result]['count'];
                               var count =  ++data;
                               arr1[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr1.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr1);
                           if(result != "false")
                           {
                               var count = arr1[result]['count'];
                               arr1[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr1.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr1.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Panorama Hills" && commonlobbyActivitiesValue[i].tower == "2B")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr2.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr2);
                           
                           if(result != "false")
                           {
                               var data = arr2[result]['count'];
                               var count =  ++data;
                               arr2[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr2.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr2);
                           if(result != "false")
                           {
                               var count = arr2[result]['count'];
                               arr2[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr2.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr2.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Panorama Hills" && commonlobbyActivitiesValue[i].tower == "2C")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr3.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr3);
                           
                           if(result != "false")
                           {
                               var data = arr3[result]['count'];
                               var count =  ++data;
                               arr3[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr3.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr3);
                           if(result != "false")
                           {
                               var count = arr2[result]['count'];
                               arr3[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr3.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr3.push(obj);
                       }
                   }
              });
            }          
            else if(commonlobbyActivitiesValue[i].project_name =="Panorama Hills" && commonlobbyActivitiesValue[i].tower == "2D")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr4.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr4);
                           
                           if(result != "false")
                           {
                               var data = arr4[result]['count'];
                               var count =  ++data;
                               arr4[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr4.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr4);
                           if(result != "false")
                           {
                               var count = arr4[result]['count'];
                               arr4[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr4.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr4.push(obj);
                       }
                   }
                   commonlobbyExcelVizag(floor_count,arr1,arr2,arr3,arr4);
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Greenfield Phase 1" && commonlobbyActivitiesValue[i].tower == "Tower A")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
         
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr6.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr6);
                           
                           if(result != "false")
                           {
                               var data = arr6[result]['count'];
                               var count =  ++data;
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr6);
                           if(result != "false")
                           {
                               var count = arr6[result]['count'];
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr6.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Greenfield Phase 1" && commonlobbyActivitiesValue[i].tower == "Tower B")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();

                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr7.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr7);
                           
                           if(result != "false")
                           {
                               var data = arr7[result]['count'];
                               var count =  ++data;
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr7);
                           if(result != "false")
                           {
                               var count = arr7[result]['count'];
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr7.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Greenfield Phase 1" && commonlobbyActivitiesValue[i].tower == "Tower C")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr8.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr8);
                           
                           if(result != "false")
                           {
                               var data = arr8[result]['count'];
                               var count =  ++data;
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr8);
                           if(result != "false")
                           {
                               var count = arr8[result]['count'];
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr8.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Greenfield Phase 1" && commonlobbyActivitiesValue[i].tower == "Tower D")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr9.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr9);
                           
                           if(result != "false")
                           {
                               var data = arr9[result]['count'];
                               var count =  ++data;
                               arr9[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr9.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr9);
                           if(result != "false")
                           {
                               var count = arr9[result]['count'];
                               arr9[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr9.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr9.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Greenfield Phase 1" && commonlobbyActivitiesValue[i].tower == "Tower E")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr10.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr10);
                           
                           if(result != "false")
                           {
                               var data = arr10[result]['count'];
                               var count =  ++data;
                               arr10[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr10.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr10);
                           if(result != "false")
                           {
                               var count = arr10[result]['count'];
                               arr10[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr10.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr10.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Greenfield Phase 1" && commonlobbyActivitiesValue[i].tower == "Tower F")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr11.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr11);
                           
                           if(result != "false")
                           {
                               var data = arr11[result]['count'];
                               var count =  ++data;
                               arr11[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr11.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr11);
                           if(result != "false")
                           {
                               var count = arr11[result]['count'];
                               arr11[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr11.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr11.push(obj);
                       }
                   }
                   commonlobbyExcelGreenfield1(floor_count,arr6,arr7,arr8,arr9,arr10,arr11);
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Greenfield Phase 2" && commonlobbyActivitiesValue[i].tower == "Tower G")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr12.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr12);
                           
                           if(result != "false")
                           {
                               var data = arr12[result]['count'];
                               var count =  ++data;
                               arr12[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr12.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr12);
                           if(result != "false")
                           {
                               var count = arr12[result]['count'];
                               arr12[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr12.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr12.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Greenfield Phase 2" && commonlobbyActivitiesValue[i].tower == "Tower H")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr13.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr13);
                           
                           if(result != "false")
                           {
                               var data = arr13[result]['count'];
                               var count =  ++data;
                               arr13[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr13.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr13);
                           if(result != "false")
                           {
                               var count = arr13[result]['count'];
                               arr13[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr13.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr13.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Greenfield Phase 2" && commonlobbyActivitiesValue[i].tower == "Tower J")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr14.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr14);
                           
                           if(result != "false")
                           {
                               var data = arr14[result]['count'];
                               var count =  ++data;
                               arr14[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr14.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr14);
                           if(result != "false")
                           {
                               var count = arr14[result]['count'];
                               arr14[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr14.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr14.push(obj);
                       }
                   }
                   commonlobbyExcelGreenfield2(floor_count,arr12,arr13,arr14);
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Southern Crest" && commonlobbyActivitiesValue[i].tower == "Tower A")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr15.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr15);
                           
                           if(result != "false")
                           {
                               var data = arr15[result]['count'];
                               var count =  ++data;
                               arr15[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr15.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr15);
                           if(result != "false")
                           {
                               var count = arr15[result]['count'];
                               arr15[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr15.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr15.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Southern Crest" && commonlobbyActivitiesValue[i].tower == "Tower B")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr16.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr16);
                           
                           if(result != "false")
                           {
                               var data = arr16[result]['count'];
                               var count =  ++data;
                               arr16[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr16.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr16);
                           if(result != "false")
                           {
                               var count = arr16[result]['count'];
                               arr16[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr16.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr16.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Southern Crest" && commonlobbyActivitiesValue[i].tower == "Tower C")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr17.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr17);
                           
                           if(result != "false")
                           {
                               var data = arr17[result]['count'];
                               var count =  ++data;
                               arr17[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr17.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr17);
                           if(result != "false")
                           {
                               var count = arr17[result]['count'];
                               arr17[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr17.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr17.push(obj);
                       }
                   }
                   commonlobbyExcelSC(floor_count,arr15,arr16,arr17)
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Luxor" && commonlobbyActivitiesValue[i].tower == "Tower A")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr18.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr18);
                           
                           if(result != "false")
                           {
                               var data = arr18[result]['count'];
                               var count =  ++data;
                               arr18[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr18.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr18);
                           if(result != "false")
                           {
                               var count = arr18[result]['count'];
                               arr18[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr18.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr18.push(obj);
                       }
                   }
              });
            }   
            else if(commonlobbyActivitiesValue[i].project_name =="Luxor" && commonlobbyActivitiesValue[i].tower == "Tower B")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr19.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr19);
                           
                           if(result != "false")
                           {
                               var data = arr19[result]['count'];
                               var count =  ++data;
                               arr19[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr19.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr19);
                           if(result != "false")
                           {
                               var count = arr19[result]['count'];
                               arr19[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr19.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr19.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Luxor" && commonlobbyActivitiesValue[i].tower == "Tower C")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr20.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr20);
                           
                           if(result != "false")
                           {
                               var data = arr20[result]['count'];
                               var count =  ++data;
                               arr20[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr20.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr20);
                           if(result != "false")
                           {
                               var count = arr20[result]['count'];
                               arr20[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr20.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr20.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Luxor" && commonlobbyActivitiesValue[i].tower == "Tower D")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr21.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr21);
                           
                           if(result != "false")
                           {
                               var data = arr21[result]['count'];
                               var count =  ++data;
                               arr21[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr21.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr21);
                           if(result != "false")
                           {
                               var count = arr21[result]['count'];
                               arr21[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr21.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr21.push(obj);
                       }
                   }
              });
            }   
            else if(commonlobbyActivitiesValue[i].project_name =="Luxor" && commonlobbyActivitiesValue[i].tower == "Tower E")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr22.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr22);
                           
                           if(result != "false")
                           {
                               var data = arr22[result]['count'];
                               var count =  ++data;
                               arr22[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr22.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr22);
                           if(result != "false")
                           {
                               var count = arr22[result]['count'];
                               arr22[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr22.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr22.push(obj);
                       }
                   }
              });
            }
            else if(commonlobbyActivitiesValue[i].project_name =="Luxor" && commonlobbyActivitiesValue[i].tower == "Tower F")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr23.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr23);
                           
                           if(result != "false")
                           {
                               var data = arr23[result]['count'];
                               var count =  ++data;
                               arr23[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr23.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr23);
                           if(result != "false")
                           {
                               var count = arr23[result]['count'];
                               arr23[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr23.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr23.push(obj);
                       }
                   }
              });
            }   
            else if(commonlobbyActivitiesValue[i].project_name =="Luxor" && commonlobbyActivitiesValue[i].tower == "Tower G")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr24.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr24);
                           
                           if(result != "false")
                           {
                               var data = arr24[result]['count'];
                               var count =  ++data;
                               arr24[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr24.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr24);
                           if(result != "false")
                           {
                               var count = arr24[result]['count'];
                               arr24[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr24.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr24.push(obj);
                       }
                   }
              });
            } 
            else if(commonlobbyActivitiesValue[i].project_name =="Luxor" && commonlobbyActivitiesValue[i].tower == "Tower H")
            {
                const activitiesNode = firebase.database().ref("commonlobby_activities/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr25.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr25);
                           
                           if(result != "false")
                           {
                               var data = arr25[result]['count'];
                               var count =  ++data;
                               arr25[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr25.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr25);
                           if(result != "false")
                           {
                               var count = arr25[result]['count'];
                               arr25[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr25.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr25.push(obj);
                       }
                   }
                   commonlobbyExcelLuxor(floor_count,arr18,arr19,arr20,arr21,arr22,arr23,arr24,arr25)
              });
            }                
        }
        
        
        res.status(200).send({
            error :false,
            Array :floor_count,
        });
    
    
    }).catch(error => {
        res.status(200).send({
            error   : true,
            message : error
        });
    });
});


//api for commonloby activities
router.post('/flatfinishactivitiesPanoramahills',(req,res, next) =>{

    const dbRef = firebase.database().ref("flatfinish_PanoramaHills");

    //array to keep count of number of floors
    var floor_count = new Array();

    //to keep track of count
    var Tower2A = 0;
    var Tower2B = 0;
    var Tower2C = 0;    
    var Tower2D = 0;    
    
    
    //array for panorama hills
    var arr1 = new Array();
    var arr2 = new Array();
    var arr3 = new Array();
    var arr4 = new Array();


    dbRef.once("value").then(snap =>{
        var flatfinishActivitiesValue = snap.val();
        var length = flatfinishActivitiesValue.length;
        for(var i=1;i<flatfinishActivitiesValue.length;i++)
        {
            var obj = {};

            if(flatfinishActivitiesValue[i].project_name == "Panorama Hills")
            {
               if(flatfinishActivitiesValue[i].tower == "2A")
               {
                    Tower2A++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    if(floor_count.length != 0)
                    {
                      var result = checkObject(tower,floor_count,Tower2A);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = Tower2A;
                      }
                    }
                    else
                    {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = Tower2A;
                       floor_count.push(obj);
                    }
               }
               else if(flatfinishActivitiesValue[i].tower == "2B")
               {
                     Tower2B++;
                     var tower  = flatfinishActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,Tower2B);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = Tower2B;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = Tower2B;
                      floor_count.push(obj);
                     }
                }
                else if(flatfinishActivitiesValue[i].tower == "2C")
                {
                    Tower2C++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    var result = checkObject(tower,floor_count,Tower2C);
                    if(result != "false")
                    {
                      floor_count[result]['total_unit'] = Tower2C;
                    }
                    else
                    {
                     obj['tower_name'] = tower,
                     obj['total_unit'] = Tower2C;
                     floor_count.push(obj);
                    }
               }
               else if(flatfinishActivitiesValue[i].tower == "2D")
               {
                   Tower2D++;
                   var tower  = flatfinishActivitiesValue[i].tower;
                   var result = checkObject(tower,floor_count,Tower2D);
                   if(result != "false")
                   {
                     floor_count[result]['total_unit'] = Tower2D;
                   }
                   else
                   {
                    obj['tower_name'] = tower,
                    obj['total_unit'] = Tower2D;
                    floor_count.push(obj);
                   }
              }
            }

            //to access activities
            if(flatfinishActivitiesValue[i].project_name =="Panorama Hills" && flatfinishActivitiesValue[i].tower == "2A")
            {  
                const activitiesNode = firebase.database().ref("flatfinish_PanoramaHills/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr1.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr1);
                           
                           if(result != "false")
                           {
                               var data = arr1[result]['count'];
                               var count =  ++data;
                               arr1[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr1.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr1);
                           if(result != "false")
                           {
                               var count = arr1[result]['count'];
                               arr1[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr1.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr1.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Panorama Hills" && flatfinishActivitiesValue[i].tower == "2B")
            {
                const activitiesNode = firebase.database().ref("flatfinish_PanoramaHills/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr2.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr2);
                           
                           if(result != "false")
                           {
                               var data = arr2[result]['count'];
                               var count =  ++data;
                               arr2[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr2.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr2);
                           if(result != "false")
                           {
                               var count = arr2[result]['count'];
                               arr2[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr2.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr2.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Panorama Hills" && flatfinishActivitiesValue[i].tower == "2C")
            {
                const activitiesNode = firebase.database().ref("flatfinish_PanoramaHills/"+i+"/activities");
         
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr3.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr3);
                           
                           if(result != "false")
                           {
                               var data = arr3[result]['count'];
                               var count =  ++data;
                               arr3[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr3.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr3);
                           if(result != "false")
                           {
                               var count = arr3[result]['count'];
                               arr3[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr3.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr3.push(obj);
                       }
                   }
              });
            }          
            else if(flatfinishActivitiesValue[i].project_name =="Panorama Hills" && flatfinishActivitiesValue[i].tower == "2D")
            {
                const activitiesNode = firebase.database().ref("flatfinish_PanoramaHills/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr4.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr4);
                           
                           if(result != "false")
                           {
                               var data = arr4[result]['count'];
                               var count =  ++data;
                               arr4[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr4.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr4);
                           if(result != "false")
                           {
                               var count = arr4[result]['count'];
                               arr4[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr4.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr4.push(obj);
                       }
                   }
                   flatfinishExcelVizag(floor_count,arr1,arr2,arr3,arr4);
              });
            }
       } 
    
       res.status(200).send({
        error :false,
        message :"Excel file generated",
    });


    }).catch(error => {
        res.status(200).send({
            error   : true,
            message : error
        });
    });
});

//api for flat finsh greenfield phase1
router.post('/flatfinishactivitiesGfphase1',(req,res, next) =>{

    const dbRef = firebase.database().ref("flatfinish_GreenfieldPhase1");

    //array to keep count of number of floors
    var floor_count = new Array();

    //greenfield phase1
    var TowerA = 0;
    var TowerB = 0;
    var TowerC = 0;
    var TowerD = 0;
    var TowerE = 0;
    var TowerF = 0;
    
    
    //arr for greenphase-1 
    var arr6 = new Array();//greenphase-1  Tower A
    var arr7 = new Array();//greenphase-1  Tower B
    var arr8 = new Array();//greenphase-1  Tower C
    var arr9 = new Array();//greenphase-1  Tower D
    var arr10 = new Array();//greenphase-1  Tower E
    var arr11 = new Array();//greenphase-1  Tower F


    dbRef.once("value").then(snap =>{
        var flatfinishActivitiesValue = snap.val();
        var length = flatfinishActivitiesValue.length;

        for(var i=1;i<flatfinishActivitiesValue.length;i++)
        {
            var obj = {};

            if(flatfinishActivitiesValue[i].project_name == "Greenfield Phase 1")
            {
               if(flatfinishActivitiesValue[i].tower == "Tower A")
               {
                    TowerA++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    if(floor_count.length != 0)
                    {
                      var result = checkObject(tower,floor_count,TowerA);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = TowerA;
                      }
                    }
                    else
                    {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = TowerA;
                       floor_count.push(obj);
                    }
               }
               else if(flatfinishActivitiesValue[i].tower == "Tower B")
               {
                     TowerB++;
                     var tower  = flatfinishActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerB);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = TowerB;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = TowerB;
                      floor_count.push(obj);
                     }
                }
                else if(flatfinishActivitiesValue[i].tower == "Tower C")
                {
                    TowerC++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    var result = checkObject(tower,floor_count,TowerC);
                    if(result != "false")
                    {
                      floor_count[result]['total_unit'] = TowerC;
                    }
                    else
                    {
                     obj['tower_name'] = tower,
                     obj['total_unit'] = TowerC;
                     floor_count.push(obj);
                    }
               }
               else if(flatfinishActivitiesValue[i].tower == "Tower D")
               {
                   TowerD++;
                   var tower  = flatfinishActivitiesValue[i].tower;
                   var result = checkObject(tower,floor_count,TowerD);
                   if(result != "false")
                   {
                     floor_count[result]['total_unit'] = TowerD;
                   }
                   else
                   {
                    obj['tower_name'] = tower,
                    obj['total_unit'] = TowerD;
                    floor_count.push(obj);
                   }
              }
              else if(flatfinishActivitiesValue[i].tower == "Tower E")
              {
                  TowerE++;
                  var tower  = flatfinishActivitiesValue[i].tower;
                  var result = checkObject(tower,floor_count,TowerE);
                  if(result != "false")
                  {
                    floor_count[result]['total_unit'] = TowerE;
                  }
                  else
                  {
                   obj['tower_name'] = tower,
                   obj['total_unit'] = TowerE;
                   floor_count.push(obj);
                  }
             }
             else if(flatfinishActivitiesValue[i].tower == "Tower F")
             {
                TowerF++;
                 var tower  = flatfinishActivitiesValue[i].tower;
                 var result = checkObject(tower,floor_count,TowerF);
                 if(result != "false")
                 {
                   floor_count[result]['total_unit'] = TowerF;
                 }
                 else
                 {
                  obj['tower_name'] = tower,
                  obj['total_unit'] = TowerF;
                  floor_count.push(obj);
                 }
              }
            }

            //to access activities
            if(flatfinishActivitiesValue[i].project_name =="Greenfield Phase 1" && flatfinishActivitiesValue[i].tower == "Tower A")
            {  
                const activitiesNode = firebase.database().ref("flatfinish_GreenfieldPhase1/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr6.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr6);
                           
                           if(result != "false")
                           {
                               var data = arr6[result]['count'];
                               var count =  ++data;
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr6);
                           if(result != "false")
                           {
                               var count = arr6[result]['count'];
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr6.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Greenfield Phase 1" && flatfinishActivitiesValue[i].tower == "Tower B")
            {
                const activitiesNode = firebase.database().ref("flatfinish_GreenfieldPhase1/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr7.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr7);
                           
                           if(result != "false")
                           {
                               var data = arr7[result]['count'];
                               var count =  ++data;
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr7);
                           if(result != "false")
                           {
                               var count = arr7[result]['count'];
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr7.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Greenfield Phase 1" && flatfinishActivitiesValue[i].tower == "Tower C")
            {
                const activitiesNode = firebase.database().ref("flatfinish_GreenfieldPhase1/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr8.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr8);
                           
                           if(result != "false")
                           {
                               var data = arr8[result]['count'];
                               var count =  ++data;
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr8);
                           if(result != "false")
                           {
                               var count = arr8[result]['count'];
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr8.push(obj);
                       }
                   }
              });
            }          
            else if(flatfinishActivitiesValue[i].project_name =="Greenfield Phase 1" && flatfinishActivitiesValue[i].tower == "Tower D")
            {
                const activitiesNode = firebase.database().ref("flatfinish_GreenfieldPhase1/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr9.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr9);
                           
                           if(result != "false")
                           {
                               var data = arr9[result]['count'];
                               var count =  ++data;
                               arr9[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr9.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr9);
                           if(result != "false")
                           {
                               var count = arr9[result]['count'];
                               arr9[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr9.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr9.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Greenfield Phase 1" && flatfinishActivitiesValue[i].tower == "Tower E")
            {
                const activitiesNode = firebase.database().ref("flatfinish_GreenfieldPhase1/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr10.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr10);
                           
                           if(result != "false")
                           {
                               var data = arr10[result]['count'];
                               var count =  ++data;
                               arr10[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr10.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr10);
                           if(result != "false")
                           {
                               var count = arr10[result]['count'];
                               arr10[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr10.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr10.push(obj);
                       }
                   }
                });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Greenfield Phase 1" && flatfinishActivitiesValue[i].tower == "Tower F")
            {
                const activitiesNode = firebase.database().ref("flatfinish_GreenfieldPhase1/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr11.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr11);
                           
                           if(result != "false")
                           {
                               var data = arr11[result]['count'];
                               var count =  ++data;
                               arr11[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr11.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr11);
                           if(result != "false")
                           {
                               var count = arr11[result]['count'];
                               arr11[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr11.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr11.push(obj);
                       }
                   }
                   //calling excel function for flatfinish activities  
                   flatfinishExcelGreenfield1(floor_count,arr6,arr7,arr8,arr9,arr10,arr11) 
                });
            }

       } 
    
    }).catch(error => {
        res.status(200).send({
            error   : true,
            message : error
        });
    });
});



//api for flat finsh greenfield phase2
router.post('/flatfinishactivitiesGfphase2',(req,res, next) =>{

    const dbRef = firebase.database().ref("flatfinish_GreenfieldPhase2");

    //array to keep count of number of floors
    var floor_count = new Array();

    //greenfield phase1
    var TowerA = 0;
    var TowerB = 0;
    var TowerC = 0;
    
    
    //arr for greenphase-1 
    var arr6 = new Array();//greenphase-1  Tower A
    var arr7 = new Array();//greenphase-1  Tower B
    var arr8 = new Array();//greenphase-1  Tower C


    dbRef.once("value").then(snap =>{
        var flatfinishActivitiesValue = snap.val();
        var length = flatfinishActivitiesValue.length;
        for(var i=1;i<flatfinishActivitiesValue.length;i++)
        {
            var obj = {};

            if(flatfinishActivitiesValue[i].project_name == "Greenfield Phase 2")
            {
               if(flatfinishActivitiesValue[i].tower == "Tower G")
               {
                    TowerA++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    if(floor_count.length != 0)
                    {
                      var result = checkObject(tower,floor_count,TowerA);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = TowerA;
                      }
                    }
                    else
                    {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = TowerA;
                       floor_count.push(obj);
                    }
               }
               else if(flatfinishActivitiesValue[i].tower == "Tower H")
               {
                     TowerB++;
                     var tower  = flatfinishActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerB);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = TowerB;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = TowerB;
                      floor_count.push(obj);
                     }
                }
                else if(flatfinishActivitiesValue[i].tower == "Tower J")
                {
                    TowerC++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    var result = checkObject(tower,floor_count,TowerC);
                    if(result != "false")
                    {
                      floor_count[result]['total_unit'] = TowerC;
                    }
                    else
                    {
                     obj['tower_name'] = tower,
                     obj['total_unit'] = TowerC;
                     floor_count.push(obj);
                    }
               }
            }

            //to access activities
            if(flatfinishActivitiesValue[i].project_name =="Greenfield Phase 2" && flatfinishActivitiesValue[i].tower == "Tower G")
            {  
                const activitiesNode = firebase.database().ref("flatfinish_GreenfieldPhase2/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr6.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr6);
                           
                           if(result != "false")
                           {
                               var data = arr6[result]['count'];
                               var count =  ++data;
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr6);
                           if(result != "false")
                           {
                               var count = arr6[result]['count'];
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr6.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Greenfield Phase 2" && flatfinishActivitiesValue[i].tower == "Tower H")
            {
                const activitiesNode = firebase.database().ref("flatfinish_GreenfieldPhase2/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr7.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr7);
                           
                           if(result != "false")
                           {
                               var data = arr7[result]['count'];
                               var count =  ++data;
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr7);
                           if(result != "false")
                           {
                               var count = arr7[result]['count'];
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr7.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Greenfield Phase 2" && flatfinishActivitiesValue[i].tower == "Tower J")
            {
                const activitiesNode = firebase.database().ref("flatfinish_GreenfieldPhase2/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr8.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr8);
                           
                           if(result != "false")
                           {
                               var data = arr8[result]['count'];
                               var count =  ++data;
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr8);
                           if(result != "false")
                           {
                               var count = arr8[result]['count'];
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr8.push(obj);
                       }
                   }
                   flatfinishExcelGreenfield2(floor_count,arr6,arr7,arr8) 

              });
            }          

       } 
    
    }).catch(error => {
        res.status(200).send({
            error   : true,
            message : error
        });
    });
});

//api for flat finsh SouthrenCrest
router.post('/flatfinishactivitiesSouthrenCrest',(req,res, next) =>{

    const dbRef = firebase.database().ref("flatfinish_SouthernCrest");

    //array to keep count of number of floors
    var floor_count = new Array();

    //Southrencrest 
    var TowerA = 0;
    var TowerB = 0;
    var TowerC = 0;
    
    
    //array for Southrencrest 
    var arr6 = new Array();//Southrencrest  Tower A
    var arr7 = new Array();//Southrencrest  Tower B
    var arr8 = new Array();//Southrencrest  Tower C


    dbRef.once("value").then(snap =>{
        var flatfinishActivitiesValue = snap.val();
        var length = flatfinishActivitiesValue.length;
        for(var i=1;i<flatfinishActivitiesValue.length;i++)
        {
            var obj = {};

            if(flatfinishActivitiesValue[i].project_name == "Southern Crest")
            {
               if(flatfinishActivitiesValue[i].tower == "Tower A")
               {
                    TowerA++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    if(floor_count.length != 0)
                    {
                      var result = checkObject(tower,floor_count,TowerA);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = TowerA;
                      }
                    }
                    else
                    {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = TowerA;
                       floor_count.push(obj);
                    }
               }
               else if(flatfinishActivitiesValue[i].tower == "Tower B")
               {
                     TowerB++;
                     var tower  = flatfinishActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerB);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = TowerB;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = TowerB;
                      floor_count.push(obj);
                     }
                }
                else if(flatfinishActivitiesValue[i].tower == "Tower C")
                {
                    TowerC++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    var result = checkObject(tower,floor_count,TowerC);
                    if(result != "false")
                    {
                      floor_count[result]['total_unit'] = TowerC;
                    }
                    else
                    {
                     obj['tower_name'] = tower,
                     obj['total_unit'] = TowerC;
                     floor_count.push(obj);
                    }
               }
            }

            //to access activities
            if(flatfinishActivitiesValue[i].project_name =="Southern Crest" && flatfinishActivitiesValue[i].tower == "Tower A")
            {  
                const activitiesNode = firebase.database().ref("flatfinish_SouthernCrest/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr6.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr6);
                           
                           if(result != "false")
                           {
                               var data = arr6[result]['count'];
                               var count =  ++data;
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr6);
                           if(result != "false")
                           {
                               var count = arr6[result]['count'];
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr6.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Southern Crest" && flatfinishActivitiesValue[i].tower == "Tower B")
            {
                const activitiesNode = firebase.database().ref("flatfinish_SouthernCrest/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr7.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr7);
                           
                           if(result != "false")
                           {
                               var data = arr7[result]['count'];
                               var count =  ++data;
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr7);
                           if(result != "false")
                           {
                               var count = arr7[result]['count'];
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr7.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Southern Crest" && flatfinishActivitiesValue[i].tower == "Tower C")
            {
                const activitiesNode = firebase.database().ref("flatfinish_SouthernCrest/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr8.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr8);
                           
                           if(result != "false")
                           {
                               var data = arr8[result]['count'];
                               var count =  ++data;
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr8);
                           if(result != "false")
                           {
                               var count = arr8[result]['count'];
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr8.push(obj);
                       }
                   }
                   flatfinishExcelSouthrenCrest(floor_count,arr6,arr7,arr8) 

              });
            }          

       } 
    
    }).catch(error => {
        res.status(200).send({
            error   : true,
            message : error
        });
    });
});


//api for flat finsh Luxor
router.post('/flatfinishactivitiesLuxor',(req,res, next) =>{

    const dbRef = firebase.database().ref("flatfinish_Luxor");

    //array to keep count of number of floors
    var floor_count = new Array();

    //Luxor
    var TowerA = 0;
    var TowerB = 0;
    var TowerC = 0;
    var TowerD = 0;
    var TowerE = 0;
    var TowerF = 0;
    var TowerG = 0;
    var TowerH = 0;
    
    
    //arr for greenphase-1 
    var arr6 = new Array();//  Tower A
    var arr7 = new Array();//  Tower B
    var arr8 = new Array();//  Tower C
    var arr9 = new Array();//  Tower D
    var arr10 = new Array();//  Tower E
    var arr11 = new Array();//  Tower F
    var arr12 = new Array();//  Tower G
    var arr13 = new Array();//  Tower H


    dbRef.once("value").then(snap =>{
        var flatfinishActivitiesValue = snap.val();
        var length = flatfinishActivitiesValue.length;

        for(var i=1;i<flatfinishActivitiesValue.length;i++)
        {
            var obj = {};

            if(flatfinishActivitiesValue[i].project_name == "Luxor")
            {
               if(flatfinishActivitiesValue[i].tower == "Tower A")
               {
                    TowerA++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    if(floor_count.length != 0)
                    {
                      var result = checkObject(tower,floor_count,TowerA);
                      if(result != "false")
                      {
                        floor_count[result]['total_unit'] = TowerA;
                      }
                    }
                    else
                    {
                       obj['tower_name'] = tower,
                       obj['total_unit'] = TowerA;
                       floor_count.push(obj);
                    }
               }
               else if(flatfinishActivitiesValue[i].tower == "Tower B")
               {
                     TowerB++;
                     var tower  = flatfinishActivitiesValue[i].tower;
                     var result = checkObject(tower,floor_count,TowerB);
                     if(result != "false")
                     {
                       floor_count[result]['total_unit'] = TowerB;
                     }
                     else
                     {
                      obj['tower_name'] = tower,
                      obj['total_unit'] = TowerB;
                      floor_count.push(obj);
                     }
                }
                else if(flatfinishActivitiesValue[i].tower == "Tower C")
                {
                    TowerC++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    var result = checkObject(tower,floor_count,TowerC);
                    if(result != "false")
                    {
                      floor_count[result]['total_unit'] = TowerC;
                    }
                    else
                    {
                     obj['tower_name'] = tower,
                     obj['total_unit'] = TowerC;
                     floor_count.push(obj);
                    }
               }
               else if(flatfinishActivitiesValue[i].tower == "Tower D")
               {
                   TowerD++;
                   var tower  = flatfinishActivitiesValue[i].tower;
                   var result = checkObject(tower,floor_count,TowerD);
                   if(result != "false")
                   {
                     floor_count[result]['total_unit'] = TowerD;
                   }
                   else
                   {
                    obj['tower_name'] = tower,
                    obj['total_unit'] = TowerD;
                    floor_count.push(obj);
                   }
              }
              else if(flatfinishActivitiesValue[i].tower == "Tower E")
              {
                  TowerE++;
                  var tower  = flatfinishActivitiesValue[i].tower;
                  var result = checkObject(tower,floor_count,TowerE);
                  if(result != "false")
                  {
                    floor_count[result]['total_unit'] = TowerE;
                  }
                  else
                  {
                   obj['tower_name'] = tower,
                   obj['total_unit'] = TowerE;
                   floor_count.push(obj);
                  }
             }
             else if(flatfinishActivitiesValue[i].tower == "Tower F")
             {
                 TowerF++;
                 var tower  = flatfinishActivitiesValue[i].tower;
                 var result = checkObject(tower,floor_count,TowerF);
                 if(result != "false")
                 {
                   floor_count[result]['total_unit'] = TowerF;
                 }
                 else
                 {
                  obj['tower_name'] = tower,
                  obj['total_unit'] = TowerF;
                  floor_count.push(obj);
                 }
              }
              else if(flatfinishActivitiesValue[i].tower == "Tower G")
              {
                  TowerG++;
                  var tower  = flatfinishActivitiesValue[i].tower;
                  var result = checkObject(tower,floor_count,TowerG);
                  if(result != "false")
                  {
                    floor_count[result]['total_unit'] = TowerG;
                  }
                  else
                  {
                   obj['tower_name'] = tower,
                   obj['total_unit'] = TowerG;
                   floor_count.push(obj);
                  }
                }
                else if(flatfinishActivitiesValue[i].tower == "Tower H")
                {
                    TowerH++;
                    var tower  = flatfinishActivitiesValue[i].tower;
                    var result = checkObject(tower,floor_count,TowerH);
                    if(result != "false")
                    {
                      floor_count[result]['total_unit'] = TowerH;
                    }
                    else
                    {
                     obj['tower_name'] = tower,
                     obj['total_unit'] = TowerH;
                     floor_count.push(obj);
                    }
                  }
            }

            //to access activities
            if(flatfinishActivitiesValue[i].project_name =="Luxor" && flatfinishActivitiesValue[i].tower == "Tower A")
            {  
                const activitiesNode = firebase.database().ref("flatfinish_Luxor/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr6.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr6);
                           
                           if(result != "false")
                           {
                               var data = arr6[result]['count'];
                               var count =  ++data;
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr6);
                           if(result != "false")
                           {
                               var count = arr6[result]['count'];
                               arr6[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr6.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr6.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Luxor" && flatfinishActivitiesValue[i].tower == "Tower B")
            {
                const activitiesNode = firebase.database().ref("flatfinish_Luxor/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr7.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr7);
                           
                           if(result != "false")
                           {
                               var data = arr7[result]['count'];
                               var count =  ++data;
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr7);
                           if(result != "false")
                           {
                               var count = arr7[result]['count'];
                               arr7[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr7.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr7.push(obj);
                       }
                   }
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Luxor" && flatfinishActivitiesValue[i].tower == "Tower C")
            {
                const activitiesNode = firebase.database().ref("flatfinish_Luxor/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
         
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr8.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr8);
                           
                           if(result != "false")
                           {
                               var data = arr8[result]['count'];
                               var count =  ++data;
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr8);
                           if(result != "false")
                           {
                               var count = arr8[result]['count'];
                               arr8[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr8.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr8.push(obj);
                       }
                   }
                   
              });
            }
            else if(flatfinishActivitiesValue[i].project_name =="Luxor" && flatfinishActivitiesValue[i].tower == "Tower D")
            {
                const activitiesNode = firebase.database().ref("flatfinish_Luxor/"+i+"/activities");
          
              //child node function
              activitiesNode.once("value").then( activityChild => {
                   var activitiesVal = activityChild.val();
          
                   //loop thru activities
                   for(var j=1;j<activitiesVal.length;j++)
                   {          
                       var obj = {};
                       if(arr9.length != 0)
                       {
                        var activityName = activitiesVal[j].activity_name;
                        if(activitiesVal[j].activity_status == "1")
                        {
                           var result = checkActivity(activityName,arr9);
                           
                           if(result != "false")
                           {
                               var data = arr9[result]['count'];
                               var count =  ++data;
                               arr9[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr9.push(obj);    
                           }
                        }
                        else if(activitiesVal[j].activity_status == "0")
                        {
                           var result = checkActivity(activityName,arr9);
                           if(result != "false")
                           {
                               var count = arr9[result]['count'];
                               arr9[result]['count'] = count;
                           }
                           else
                           {
                               obj['milestone'] =  activitiesVal[j].activity_name;
                               obj['count']     =  activitiesVal[j].activity_status;
                               arr9.push(obj);    
                           }
                        }
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr9.push(obj);
                       }
                   }

              });
            
        } 
        else if(flatfinishActivitiesValue[i].project_name =="Luxor" && flatfinishActivitiesValue[i].tower == "Tower E")
        {
            const activitiesNode = firebase.database().ref("flatfinish_Luxor/"+i+"/activities");
      
          //child node function
          activitiesNode.once("value").then( activityChild => {
               var activitiesVal = activityChild.val();
      
               //loop thru activities
               for(var j=1;j<activitiesVal.length;j++)
               {          
                   var obj = {};
                   if(arr10.length != 0)
                   {
                    var activityName = activitiesVal[j].activity_name;
                    if(activitiesVal[j].activity_status == "1")
                    {
                       var result = checkActivity(activityName,arr10);
                       
                       if(result != "false")
                       {
                           var data = arr10[result]['count'];
                           var count =  ++data;
                           arr10[result]['count'] = count;
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr10.push(obj);    
                       }
                    }
                    else if(activitiesVal[j].activity_status == "0")
                    {
                       var result = checkActivity(activityName,arr10);
                       if(result != "false")
                       {
                           var count = arr10[result]['count'];
                           arr10[result]['count'] = count;
                       }
                       else
                       {
                           obj['milestone'] =  activitiesVal[j].activity_name;
                           obj['count']     =  activitiesVal[j].activity_status;
                           arr10.push(obj);    
                       }
                    }
                   }
                   else
                   {
                       obj['milestone'] =  activitiesVal[j].activity_name;
                       obj['count']     =  activitiesVal[j].activity_status;
                       arr10.push(obj);
                   }
               }
               
         });
        
       }
       else if(flatfinishActivitiesValue[i].project_name =="Luxor" && flatfinishActivitiesValue[i].tower == "Tower F")
       {
           const activitiesNode = firebase.database().ref("flatfinish_Luxor/"+i+"/activities");
     
         //child node function
         activitiesNode.once("value").then( activityChild => {
              var activitiesVal = activityChild.val();
     
              //loop thru activities
              for(var j=1;j<activitiesVal.length;j++)
              {          
                  var obj = {};
                  if(arr11.length != 0)
                  {
                   var activityName = activitiesVal[j].activity_name;
                   if(activitiesVal[j].activity_status == "1")
                   {
                      var result = checkActivity(activityName,arr11);
                      
                      if(result != "false")
                      {
                          var data = arr11[result]['count'];
                          var count =  ++data;
                          arr11[result]['count'] = count;
                      }
                      else
                      {
                          obj['milestone'] =  activitiesVal[j].activity_name;
                          obj['count']     =  activitiesVal[j].activity_status;
                          arr11.push(obj);    
                      }
                   }
                   else if(activitiesVal[j].activity_status == "0")
                   {
                      var result = checkActivity(activityName,arr11);
                      if(result != "false")
                      {
                          var count = arr11[result]['count'];
                          arr11[result]['count'] = count;
                      }
                      else
                      {
                          obj['milestone'] =  activitiesVal[j].activity_name;
                          obj['count']     =  activitiesVal[j].activity_status;
                          arr11.push(obj);    
                      }
                   }
                  }
                  else
                  {
                      obj['milestone'] =  activitiesVal[j].activity_name;
                      obj['count']     =  activitiesVal[j].activity_status;
                      arr11.push(obj);
                  }
              }

         });
       
       }
       else if(flatfinishActivitiesValue[i].project_name =="Luxor" && flatfinishActivitiesValue[i].tower == "Tower G")
       {
           const activitiesNode = firebase.database().ref("flatfinish_Luxor/"+i+"/activities");
     
         //child node function
         activitiesNode.once("value").then( activityChild => {
              var activitiesVal = activityChild.val();
     
              //loop thru activities
              for(var j=1;j<activitiesVal.length;j++)
              {          
                  var obj = {};
                  if(arr12.length != 0)
                  {
                   var activityName = activitiesVal[j].activity_name;
                   if(activitiesVal[j].activity_status == "1")
                   {
                      var result = checkActivity(activityName,arr12);
                      
                      if(result != "false")
                      {
                          var data = arr12[result]['count'];
                          var count =  ++data;
                          arr12[result]['count'] = count;
                      }
                      else
                      {
                          obj['milestone'] =  activitiesVal[j].activity_name;
                          obj['count']     =  activitiesVal[j].activity_status;
                          arr12.push(obj);    
                      }
                   }
                   else if(activitiesVal[j].activity_status == "0")
                   {
                      var result = checkActivity(activityName,arr12);
                      if(result != "false")
                      {
                          var count = arr12[result]['count'];
                          arr12[result]['count'] = count;
                      }
                      else
                      {
                          obj['milestone'] =  activitiesVal[j].activity_name;
                          obj['count']     =  activitiesVal[j].activity_status;
                          arr12.push(obj);    
                      }
                   }
                  }
                  else
                  {
                      obj['milestone'] =  activitiesVal[j].activity_name;
                      obj['count']     =  activitiesVal[j].activity_status;
                      arr12.push(obj);
                  }
              }
         });
       
      }  
      else if(flatfinishActivitiesValue[i].project_name =="Luxor" && flatfinishActivitiesValue[i].tower == "Tower H")
      {
          const activitiesNode = firebase.database().ref("flatfinish_Luxor/"+i+"/activities");
   
        //child node function
        activitiesNode.once("value").then( activityChild => {
             var activitiesVal = activityChild.val();
    
             //loop thru activities
             for(var j=1;j<activitiesVal.length;j++)
             {          
                 var obj = {};
                 if(arr13.length != 0)
                 {
                  var activityName = activitiesVal[j].activity_name;
                  if(activitiesVal[j].activity_status == "1")
                  {
                     var result = checkActivity(activityName,arr13);
                     
                     if(result != "false")
                     {
                         var data = arr13[result]['count'];
                         var count =  ++data;
                         arr13[result]['count'] = count;
                     }
                     else
                     {
                         obj['milestone'] =  activitiesVal[j].activity_name;
                         obj['count']     =  activitiesVal[j].activity_status;
                         arr13.push(obj);    
                     }
                  }
                  else if(activitiesVal[j].activity_status == "0")
                  {
                     var result = checkActivity(activityName,arr13);
                     if(result != "false")
                     {
                         var count = arr12[result]['count'];
                         arr13[result]['count'] = count;
                     }
                     else
                     {
                         obj['milestone'] =  activitiesVal[j].activity_name;
                         obj['count']     =  activitiesVal[j].activity_status;
                         arr13.push(obj);    
                     }
                  }
                 }
                 else
                 {
                     obj['milestone'] =  activitiesVal[j].activity_name;
                     obj['count']     =  activitiesVal[j].activity_status;
                     arr13.push(obj);
                 }
             }
             
             flatfinishExcelLuxor(floor_count,arr6,arr7,arr8,arr9,arr10,arr11,arr12,arr13) 

        });
      
        }
     }
     
     res.status(200).send({
        error :false,
        message :"Excel file generated",
    });
    
    }).catch(error => {
        res.status(200).send({
            error   : true,
            message : error
        });
    });
});



//api to upload for s3 bucket
router.post('/fileupload',(req, res, next) => {

 //configuring the AWS environment
 AWS.config.update({
    accessKeyId: "AKIAIJXPK6OEO73DRBIA",
    secretAccessKey: "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"
 });

 var s3 = new AWS.S3();
 var filePath = "api/routes/StructureActivities_SouthrenCrest.xlsx";

 //configuring parameters
  var params = {
  Bucket: 'streetsmartb2',
  Body : fs.createReadStream(filePath),
  Key : "folder/"+Date.now()+"_"+path.basename(filePath)
  };

   s3.upload(params, function (err, data) {
   //handle error
   if (err) {
    console.log("Error", err);
   }

   //success
   if (data) {
    console.log("Uploaded in:", data.Location);
   }
  });
});

/*****************
 *****************
 * Function Part**
 *****************
 *****************/

//function to check weather object exists in the array
function checkObject(tower,floor_count,Tower2A)
{
    for(var i=0;i<floor_count.length;i++)
    {
        var res = (floor_count[i]["tower_name"] == tower) ? i : "false";
    }
    return res;
}


//function to check activity
function checkActivity(activityName,arr1)
{
    for(var i=0;i<arr1.length;i++)
    {
        var res = (arr1[i]["milestone"] == activityName) ? i : "false";
        if(res != "false")
        {
            return i;
        }
    }
    return res;
}

//function for Commonlift vizag
//function to generate excel for Panorama Hills
function generateCommonLiftExcelVizag(arr,arr2,arr3,arr4,arr5)
{

    //path to generate excel
    const path = require('path');
    let excel = new Excel(path.join(__dirname,'PanoramaHills.xlsx'));

    excel.writeSheet('CommonliftActivities',['Activity','Tower2A','Tower2B','Tower2C','Tower2D','% Completed'],[arr]).then(()=>{
      excel.writeRow('CommonliftActivities',1,{
        Activity: "Units_Count",
        Tower2A : arr[0]['total_unit'],
        Tower2B : arr[1]['total_unit'],
        Tower2C : arr[2]['total_unit'],
        Tower2D : arr[3]['total_unit'],
      }).then(()=>{
        excel.writeRow('CommonliftActivities',2,{
          Activity : arr2[0].milestone,
          Tower2A  : arr2[0].count,
          Tower2B  : arr3[0].count,
          Tower2C  : arr4[0].count,
          Tower2D  : arr5[0].count,
          '% Completed':(((arr2[0].count+arr3[0].count+arr4[0].count+arr5[0].count)/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100),
        }).then(()=>{
          excel.writeRow('CommonliftActivities',3,{
            Activity : arr2[1].milestone,
            Tower2A :arr2[1].count,
            Tower2B :arr3[1].count,
            Tower2C  : arr4[1].count,
            Tower2D  : arr5[1].count,
            '% Completed':(((arr2[1].count+arr3[1].count+arr4[1].count+arr5[1].count)/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100),
          }).then(()=>{
            excel.writeRow('CommonliftActivities',4,{
              Activity : arr2[2].milestone,
              Tower2A :  arr2[2].count,
              Tower2B :  arr3[2].count,
              Tower2C  : arr4[2].count,
              Tower2D  : arr5[2].count,
              '% Completed':(((arr2[2].count+arr3[2].count+arr4[2].count+arr5[2].count)/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100),
            }).then(()=>{
              excel.writeRow('CommonliftActivities',5,{
                Activity : arr2[3].milestone,
                Tower2A :arr2[3].count,
                Tower2B :arr3[3].count,
                Tower2C  : arr4[3].count,
                Tower2D  : arr5[3].count,
                '% Completed':(((arr2[3].count+arr3[3].count+arr4[3].count+arr5[3].count)/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100),
              }).then(()=>{
                excel.writeRow('CommonliftActivities',6,{
                  Activity : arr2[4].milestone, 
                  Tower2A :arr2[4].count,
                  Tower2B :arr3[4].count,
                  Tower2C  : arr4[4].count,
                  Tower2D  : arr5[4].count,
                  '% Completed':(((arr2[4].count+arr3[4].count+arr4[4].count+arr5[4].count)/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100),
                }).then(()=>{
                  excel.writeRow('CommonliftActivities',7,{
                    Activity : arr2[5].milestone,
                    Tower2A :arr2[5].count,
                    Tower2B :arr3[5].count,
                    Tower2C  : arr4[5].count,
                    Tower2D  : arr5[5].count,
                    '% Completed':(((arr2[5].count+arr3[5].count+arr4[5].count+arr5[5].count)/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100),
                  }).then(()=>{
                    excel.writeRow('CommonliftActivities',8,{
                      Activity : arr2[6].milestone,
                      Tower2A :arr2[6].count,
                      Tower2B :arr3[6].count,
                      Tower2C  : arr4[6].count,
                      Tower2D  : arr5[6].count,
                      '% Completed':(((arr2[6].count+arr3[6].count+arr4[6].count+arr5[6].count)/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100),
                    })                                                       
                  });
                });
              });
            });
          });
        });
      });
    });
    return;
}


//function to generate commonlift excel for greenphase 1
function generateCommonLiftExcelGFPhase1(arr,arr6,arr7,arr8,arr9,arr10,arr11)
{
    const path = require('path');
    let excel = new Excel(path.join(__dirname,'Greenfieldphase1.xlsx'));
    
    excel.writeSheet('CommonliftActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','CompletedPercent'],[arr]).then(()=>{
        excel.writeRow('CommonliftActivities',1,{
            Activity: "Unit Count",
            TowerA : arr[4]['total_unit'],
            TowerB : arr[5]['total_unit'],
            TowerC : arr[6]['total_unit'],
            TowerD : arr[7]['total_unit'],
            TowerE : arr[8]['total_unit'],
            TowerF : arr[9]['total_unit'],
            }).then(()=>{//1 then
                    excel.writeRow('CommonliftActivities',2,{
                        Activity : arr6[0].milestone,
                        TowerA : arr6[0]['count'],
                        TowerB : arr7[0]['count'],
                        TowerC : arr8[0]['count'],
                        TowerD : arr9[0]['count'],
                        TowerE : arr10[0]['count'],
                        TowerF : arr11[0]['count'],
                        CompletedPercent : ((arr6[0]['count']+arr7[0]['count']+arr8[0]['count']+arr9[0]['count']+arr10[0]['count']+arr11[0]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                        }).then(()=>{//2nd then
                            excel.writeRow('CommonliftActivities',3,{
                                Activity : arr6[1].milestone,
                                TowerA : arr6[1]['count'],
                                TowerB : arr7[1]['count'],
                                TowerC : arr8[1]['count'],
                                TowerD : arr9[1]['count'],
                                TowerE : arr10[1]['count'],
                                TowerF : arr11[1]['count'],
                                CompletedPercent : ((arr6[1]['count']+arr7[1]['count']+arr8[1]['count']+arr9[1]['count']+arr10[1]['count']+arr11[1]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                }).then(()=>{//3rd then
                                    excel.writeRow('CommonliftActivities',4,{
                                        Activity : arr6[2].milestone,
                                        TowerA : arr6[2]['count'],
                                        TowerB : arr7[2]['count'],
                                        TowerC : arr8[2]['count'],
                                        TowerD : arr9[2]['count'],
                                        TowerE : arr10[2]['count'],
                                        TowerF : arr11[2]['count'],
                                        CompletedPercent : ((arr6[2]['count']+arr7[2]['count']+arr8[2]['count']+arr9[2]['count']+arr10[2]['count']+arr11[2]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                        }).then(()=>{//4th then
                                            excel.writeRow('CommonliftActivities',5,{
                                                Activity : arr6[3].milestone,
                                                TowerA : arr6[3]['count'],
                                                TowerB : arr7[3]['count'],
                                                TowerC : arr8[3]['count'],
                                                TowerD : arr9[3]['count'],
                                                TowerE : arr10[3]['count'],
                                                TowerF : arr11[3]['count'],
                                                CompletedPercent : ((arr6[3]['count']+arr7[3]['count']+arr8[3]['count']+arr9[3]['count']+arr10[3]['count']+arr11[3]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                }).then(()=>{//5th then
                                                    excel.writeRow('CommonliftActivities',6,{
                                                        Activity : arr6[4].milestone,
                                                        TowerA : arr6[4]['count'],
                                                        TowerB : arr7[4]['count'],
                                                        TowerC : arr8[4]['count'],
                                                        TowerD : arr9[4]['count'],
                                                        TowerE : arr10[4]['count'],
                                                        TowerF : arr11[4]['count'],
                                                        CompletedPercent : ((arr6[4]['count']+arr7[4]['count']+arr8[4]['count']+arr9[4]['count']+arr10[4]['count']+arr11[4]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                        }).then(()=>{//8th then
                                                          excel.writeRow('CommonliftActivities',7,{
                                                              Activity : arr6[5].milestone,
                                                              TowerA : arr6[5]['count'],
                                                              TowerB : arr7[5]['count'],
                                                              TowerC : arr8[5]['count'],
                                                              TowerD : arr9[5]['count'],
                                                              TowerE : arr10[5]['count'],
                                                              TowerF : arr11[5]['count'],
                                                              CompletedPercent : ((arr6[5]['count']+arr7[5]['count']+arr8[5]['count']+arr9[5]['count']+arr10[5]['count']+arr11[5]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                              }).then(()=>{//9th then
                                                                excel.writeRow('CommonliftActivities',8,{
                                                                    Activity : arr6[6].milestone,
                                                                    TowerA : arr6[6]['count'],
                                                                    TowerB : arr7[6]['count'],
                                                                    TowerC : arr8[6]['count'],
                                                                    TowerD : arr9[6]['count'],
                                                                    TowerE : arr10[6]['count'],
                                                                    TowerF : arr11[6]['count'],
                                                                    CompletedPercent : ((arr6[6]['count']+arr7[6]['count']+arr8[6]['count']+arr9[6]['count']+arr10[6]['count']+arr11[6]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                    }).then(()=>{//11th then
                                                                      excel.writeRow('CommonliftActivities',9,{
                                                                          Activity : arr6[7].milestone,
                                                                          TowerA : arr6[7]['count'],
                                                                          TowerB : arr7[7]['count'],
                                                                          TowerC : arr8[7]['count'],
                                                                          TowerD : arr9[7]['count'],
                                                                          TowerE : arr10[7]['count'],
                                                                          TowerF : arr11[7]['count'],
                                                                          CompletedPercent : ((arr6[7]['count']+arr7[7]['count']+arr8[7]['count']+arr9[7]['count']+arr10[7]['count']+arr11[7]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                          }).then(()=>{//12th then
                                                                            excel.writeRow('CommonliftActivities',10,{
                                                                                Activity : arr6[8].milestone,
                                                                                TowerA : arr6[8]['count'],
                                                                                TowerB : arr7[8]['count'],
                                                                                TowerC : arr8[8]['count'],
                                                                                TowerD : arr9[8]['count'],
                                                                                TowerE : arr10[8]['count'],
                                                                                TowerF : arr11[8]['count'],
                                                                                CompletedPercent : ((arr6[8]['count']+arr7[8]['count']+arr8[8]['count']+arr9[8]['count']+arr10[8]['count']+arr11[8]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                }).then(()=>{//13th then
                                                                                  excel.writeRow('CommonliftActivities',11,{
                                                                                      Activity : arr6[9].milestone,
                                                                                      TowerA : arr6[9]['count'],
                                                                                      TowerB : arr7[9]['count'],
                                                                                      TowerC : arr8[9]['count'],
                                                                                      TowerD : arr9[9]['count'],
                                                                                      TowerE : arr10[9]['count'],
                                                                                      TowerF : arr11[9]['count'],
                                                                                      CompletedPercent : ((arr6[9]['count']+arr7[9]['count']+arr8[9]['count']+arr9[9]['count']+arr10[9]['count']+arr11[9]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                      })
                                                                                  });
                                                                              });                                                                                                                                                                
                                                                          });
                                                                      });
                                                                });
                                                          });
                                                    });
                                            });
                                    });
                            });  
                    });
    return;
}

//Generate Excel for Greenfield Phase2 
function generateCommonLiftExcelGFPhase2(arr,arr12,arr13,arr14)
{
    const path = require('path');
    let excel = new Excel(path.join(__dirname,'Greenfieldphase2.xlsx'));
    
    excel.writeSheet('CommonliftActivities',['Activity','TowerG','TowerH','TowerJ','CompletedPercent'],[arr]).then(()=>{
        excel.writeRow('CommonliftActivities',1,{
            Activity: "Unit Count",
            TowerG : arr[0]['total_unit'],
            TowerH : arr[1]['total_unit'],
            TowerJ : arr[2]['total_unit'],
        }).then(()=>{//1 then
                    excel.writeRow('CommonliftActivities',2,{
                        Activity : arr12[0].milestone,
                        TowerG : arr12[0]['count'],
                        TowerH : arr13[0]['count'],
                        TowerJ : arr14[0]['count'],
                        CompletedPercent : ((arr12[0]['count']+arr13[0]['count']+arr14[0]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                        }).then(()=>{//2nd then
                            excel.writeRow('CommonliftActivities',3,{
                                Activity : arr12[1].milestone,
                                TowerG : arr12[1]['count'],
                                TowerH : arr13[1]['count'],
                                TowerJ : arr14[1]['count'],
                                CompletedPercent : ((arr12[1]['count']+arr13[1]['count']+arr14[1]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                                        }).then(()=>{//3rd then
                                    excel.writeRow('CommonliftActivities',4,{
                                        Activity : arr12[2].milestone,
                                        TowerG : arr12[2]['count'],
                                        TowerH : arr13[2]['count'],
                                        TowerJ : arr14[2]['count'],
                                        CompletedPercent : ((arr12[2]['count']+arr13[2]['count']+arr14[2]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                                        }).then(()=>{//4th then
                                            excel.writeRow('CommonliftActivities',5,{
                                                Activity : arr12[3].milestone,
                                                TowerG : arr12[3]['count'],
                                                TowerH : arr13[3]['count'],
                                                TowerJ : arr14[3]['count'],
                                                CompletedPercent : ((arr12[3]['count']+arr13[3]['count']+arr14[3]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                                                  }).then(()=>{//5th then
                                                    excel.writeRow('CommonliftActivities',6,{
                                                        Activity : arr12[4].milestone,
                                                        TowerG : arr12[4]['count'],
                                                        TowerH : arr13[4]['count'],
                                                        TowerJ : arr14[4]['count'],
                                                        CompletedPercent : ((arr12[4]['count']+arr13[4]['count']+arr14[4]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                                                        }).then(()=>{//6th then
                                                            excel.writeRow('CommonliftActivities',7,{
                                                                Activity : arr12[5].milestone,
                                                                TowerG : arr12[5]['count'],
                                                                TowerH : arr13[5]['count'],
                                                                TowerJ : arr14[5]['count'],
                                                                CompletedPercent : ((arr12[5]['count']+arr13[5]['count']+arr14[5]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                                                                    }).then(()=>{//7th then
                                                                    excel.writeRow('CommonliftActivities',8,{
                                                                        Activity : arr12[6].milestone,
                                                                        TowerG : arr12[6]['count'],
                                                                        TowerH : arr13[6]['count'],
                                                                        TowerJ : arr14[6]['count'],
                                                                        CompletedPercent : ((arr12[6]['count']+arr13[6]['count']+arr14[6]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                                                                         }).then(()=>{//8th then
                                                                        excel.writeRow('CommonliftActivities',9,{
                                                                            Activity : arr12[7].milestone,
                                                                            TowerG : arr12[7]['count'],
                                                                            TowerH : arr13[7]['count'],
                                                                            TowerJ : arr14[7]['count'],
                                                                            CompletedPercent : ((arr12[7]['count']+arr13[7]['count']+arr14[7]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                                                                           }).then(()=>{//9th then
                                                                            excel.writeRow('CommonliftActivities',10,{
                                                                                Activity : arr12[8].milestone,
                                                                                TowerG : arr12[8]['count'],
                                                                                TowerH : arr13[8]['count'],
                                                                                TowerJ : arr14[8]['count'],
                                                                                CompletedPercent : ((arr12[8]['count']+arr13[8]['count']+arr14[8]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                                                                                 }).then(()=>{//10th then
                                                                                excel.writeRow('CommonliftActivities',11,{
                                                                                    Activity : arr12[9].milestone,
                                                                                    TowerG : arr12[9]['count'],
                                                                                    TowerH : arr13[9]['count'],
                                                                                    TowerJ : arr14[9]['count'],
                                                                                    CompletedPercent : ((arr12[9]['count']+arr13[9]['count']+arr14[9]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']))*100+"%",
                                                                                   })
                                                                                    });
                                                                                });
                                                                            });
                                                                        });
                                                            });
                                                    });  
                                            });
                                    });
                            });
                    });
        });                                                                                                   
    return;
}



//function to generate excel for vizag property panorama hills
function commonlobbyExcelVizag(floor_count,arr1,arr2,arr3,arr4)
{

        //path to generate excel
        const path = require('path');
        let excel = new Excel(path.join(__dirname,'PanoramaHills.xlsx'));

        excel.writeSheet('CommonLobby',['Activity','Tower2A','Tower2B','Tower2C','Tower2D','CompletedPercent'],[floor_count]).then(()=>{
            excel.writeRow('CommonLobby',1,{
               Activity: "Floor Count",
               Tower2A : floor_count[0]['total_unit'],
               Tower2B : floor_count[1]['total_unit'],
               Tower2C : floor_count[2]['total_unit'],
               Tower2D : floor_count[3]['total_unit'],
               
           }).then(()=>{//1st then
               excel.writeRow('CommonLobby',2,{
                Activity: arr1[0].milestone,
                Tower2A : arr1[0].count,
                Tower2B : arr2[0].count,
                Tower2C : arr3[0].count,
                Tower2D : arr4[0].count,
                CompletedPercent : ((arr1[0]['count']+arr2[0]['count']+arr3[0]['count']+arr4[0]['count'])/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100+"%"
             }).then(()=>{//2nd then
                excel.writeRow('CommonLobby',3,{
                    Activity: arr1[1].milestone,
                    Tower2A : arr1[1].count,
                    Tower2B : arr2[1].count,
                    Tower2C : arr3[1].count,
                    Tower2D : arr4[1].count,
                    CompletedPercent : ((arr1[1]['count']+arr2[1]['count']+arr3[1]['count']+arr4[1]['count'])/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100+"%"
                }).then(()=>{//3rd then
                    excel.writeRow('CommonLobby',4,{
                        Activity: arr1[2].milestone,
                        Tower2A : arr1[2].count,
                        Tower2B : arr2[2].count,
                        Tower2C : arr3[2].count,
                        Tower2D : arr4[2].count,
                        CompletedPercent : ((arr1[2]['count']+arr2[2]['count']+arr3[2]['count']+arr4[2]['count'])/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100+"%"
    
                    }).then(()=>{//4th then
                        excel.writeRow('CommonLobby',5,{
                            Activity: arr1[3].milestone,
                            Tower2A : arr1[3].count,
                            Tower2B : arr2[3].count,
                            Tower2C : arr3[3].count,
                            Tower2D : arr4[3].count,
                            CompletedPercent : ((arr1[3]['count']+arr2[3]['count']+arr3[3]['count']+arr4[3]['count'])/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100+"%"
                                    
                        }).then(()=>{//5th then
                            excel.writeRow('CommonLobby',6,{
                                Activity: arr1[4].milestone,
                                Tower2A : arr1[4].count,
                                Tower2B : arr2[4].count,
                                Tower2C : arr3[4].count,
                                Tower2D : arr4[4].count,  
                                CompletedPercent : ((arr1[4]['count']+arr2[4]['count']+arr3[4]['count']+arr4[4]['count'])/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100+"%"  
        
                            }).then(()=>{//6th then
                                excel.writeRow('CommonLobby',7,{
                                    Activity: arr1[5].milestone,
                                    Tower2A : arr1[5].count,
                                    Tower2B : arr2[5].count,
                                    Tower2C : arr3[5].count,
                                    Tower2D : arr4[5].count,
                                    CompletedPercent : ((arr1[5]['count']+arr2[5]['count']+arr3[5]['count']+arr4[5]['count'])/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100+"%"           
                                })
                            });//6th then ending
                        });//5th then ening
                    });//4th then ending                   
                });//3rd then ending   
             });//2nd then ending
           });//1st then ending
        });
  return;    
}


function generateCommonLiftExcelSC(arr,arr15,arr16,arr17)
{
  const path = require('path');
    let excel = new Excel(path.join(__dirname,'SouthrenCrest.xlsx'));
    
    excel.writeSheet('CommonliftActivities',['Activity','TowerA','TowerB','TowerC','CompletedPercent'],[arr]).then(()=>{
        excel.writeRow('CommonliftActivities',1,{
            Activity: "Unit Count",
            TowerA : arr[10]['total_unit'],
            TowerB : arr[11]['total_unit'],
            TowerC : arr[12]['total_unit'],
        }).then(()=>{//1 then
                    excel.writeRow('CommonliftActivities',2,{
                        Activity : arr15[0].milestone,
                        TowerA : arr15[0]['count'],
                        TowerB : arr16[0]['count'],
                        TowerC : arr17[0]['count'],
                        CompletedPercent : ((arr15[0]['count']+arr16[0]['count']+arr17[0]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                        }).then(()=>{//2nd then
                            excel.writeRow('CommonliftActivities',3,{
                                Activity : arr15[1].milestone,
                                TowerA : arr15[1]['count'],
                                TowerB : arr16[1]['count'],
                                TowerC : arr17[1]['count'],
                                CompletedPercent : ((arr15[1]['count']+arr16[1]['count']+arr17[1]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                }).then(()=>{//3rd then
                                    excel.writeRow('CommonliftActivities',4,{
                                        Activity : arr15[2].milestone,
                                        TowerA : arr15[2]['count'],
                                        TowerB : arr16[2]['count'],
                                        TowerC : arr17[2]['count'],
                                        CompletedPercent : ((arr15[2]['count']+arr16[2]['count']+arr17[2]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                        }).then(()=>{//4th then
                                            excel.writeRow('CommonliftActivities',5,{
                                                Activity : arr15[3].milestone,
                                                TowerA : arr15[3]['count'],
                                                TowerB : arr16[3]['count'],
                                                TowerC : arr17[3]['count'],
                                                CompletedPercent : ((arr15[3]['count']+arr16[3]['count']+arr17[3]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                }).then(()=>{//5th then
                                                    excel.writeRow('CommonliftActivities',6,{
                                                        Activity : arr15[4].milestone,
                                                        TowerA : arr15[4]['count'],
                                                        TowerB : arr16[4]['count'],
                                                        TowerC : arr17[4]['count'],
                                                        CompletedPercent : ((arr15[4]['count']+arr16[4]['count']+arr17[4]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                        }).then(()=>{//8th then
                                                          excel.writeRow('CommonliftActivities',7,{
                                                              Activity : arr15[5].milestone,
                                                              TowerA : arr15[5]['count'],
                                                              TowerB : arr16[5]['count'],
                                                              TowerC : arr17[5]['count'],
                                                              CompletedPercent : ((arr15[5]['count']+arr16[5]['count']+arr17[5]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                          });
                                                        });
                                                    });
                                            });
                                    });
                              });
                        });
                  });
      return;            
}


function generateCommonLiftExcelLuxor(arr,arr18,arr19,arr20,arr21,arr22,arr23,arr24,arr25)
{
  //path to generate excel
  const path = require('path');
  let excel = new Excel(path.join(__dirname,'Luxor.xlsx'));

      // var tower = arr[0]['tower_name'];
     excel.writeSheet('CommonliftActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','TowerG','TowerH','CompletedPercent'],[arr]).then(()=>{
               excel.writeRow('CommonliftActivities',1,{
                  Activity: "Unit Count",
                  TowerA : arr[13]['total_unit'],
                  TowerB : arr[14]['total_unit'],
                  TowerC : arr[15]['total_unit'],
                  TowerD : arr[16]['total_unit'],
                  TowerE : arr[17]['total_unit'],
                  TowerF : arr[18]['total_unit'],
                  TowerG : arr[19]['total_unit'],
                  TowerH : arr[20]['total_unit'],
              }).then(()=>{
                      excel.writeRow('CommonliftActivities',2,{
                          Activity : arr18[0].milestone,
                          TowerA  : arr18[0].count,
                          TowerB  : arr19[0].count,
                          TowerC  : arr20[0].count,
                          TowerD  : arr21[0].count,
                          TowerE  : arr22[0].count,
                          TowerF :  arr23[0].count,
                          TowerG :  arr24[0].count,
                          TowerH :  arr25[0].count,        
                          CompletedPercent : ((arr18[0]['count']+arr19[0]['count']+arr20[0]['count']+arr21[0]['count']+arr22[0]['count']+arr23[0]['count']+arr24[0]['count']+arr25[0]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                      }).then(()=>{
                      excel.writeRow('CommonliftActivities',3,{
                        Activity : arr18[1].milestone,
                        TowerA  : arr18[1].count,
                        TowerB  : arr19[1].count,
                        TowerC  : arr20[1].count,
                        TowerD  : arr21[1].count,
                        TowerE  : arr22[1].count,
                        TowerF :  arr23[1].count,
                        TowerG :  arr24[1].count,
                        TowerH :  arr25[1].count,        
                        CompletedPercent : ((arr18[1]['count']+arr19[1]['count']+arr20[1]['count']+arr21[1]['count']+arr22[1]['count']+arr23[1]['count']+arr24[1]['count']+arr25[1]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                      }).then(()=>{
                          excel.writeRow('CommonliftActivities',4,{
                            Activity : arr18[2].milestone,
                            TowerA  : arr18[2].count,
                            TowerB  : arr19[2].count,
                            TowerC  : arr20[2].count,
                            TowerD  : arr21[2].count,
                            TowerE  : arr22[2].count,
                            TowerF :  arr23[2].count,
                            TowerG :  arr24[2].count,
                            TowerH :  arr25[2].count,        
                            CompletedPercent : ((arr18[2]['count']+arr19[2]['count']+arr20[2]['count']+arr21[2]['count']+arr22[2]['count']+arr23[2]['count']+arr24[2]['count']+arr25[2]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                              }).then(()=>{
                              excel.writeRow('CommonliftActivities',5,{
                                Activity : arr18[3].milestone,
                                TowerA  : arr18[3].count,
                                TowerB  : arr19[3].count,
                                TowerC  : arr20[3].count,
                                TowerD  : arr21[3].count,
                                TowerE  : arr22[3].count,
                                TowerF :  arr23[3].count,
                                TowerG :  arr24[3].count,
                                TowerH :  arr25[3].count,        
                                CompletedPercent : ((arr18[3]['count']+arr19[3]['count']+arr20[3]['count']+arr21[3]['count']+arr22[3]['count']+arr23[3]['count']+arr24[3]['count']+arr25[3]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                  }).then(()=>{
                                  excel.writeRow('CommonliftActivities',6,{
                                    Activity : arr18[4].milestone,
                                    TowerA  : arr18[4].count,
                                    TowerB  : arr19[4].count,
                                    TowerC  : arr20[4].count,
                                    TowerD  : arr21[4].count,
                                    TowerE  : arr22[4].count,
                                    TowerF :  arr23[4].count,
                                    TowerG :  arr24[4].count,
                                    TowerH :  arr25[4].count,        
                                    CompletedPercent : ((arr18[4]['count']+arr19[4]['count']+arr20[4]['count']+arr21[4]['count']+arr22[4]['count']+arr23[4]['count']+arr24[4]['count']+arr25[4]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                      }).then(()=>{
                                      excel.writeRow('CommonliftActivities',7,{
                                        Activity : arr18[5].milestone,
                                        TowerA  : arr18[5].count,
                                        TowerB  : arr19[5].count,
                                        TowerC  : arr20[5].count,
                                        TowerD  : arr21[5].count,
                                        TowerE  : arr22[5].count,
                                        TowerF :  arr23[5].count,
                                        TowerG :  arr24[5].count,
                                        TowerH :  arr25[5].count,        
                                        CompletedPercent : ((arr18[5]['count']+arr19[5]['count']+arr20[5]['count']+arr21[5]['count']+arr22[5]['count']+arr23[5]['count']+arr24[5]['count']+arr25[5]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                            }).then(()=>{
                                          excel.writeRow('CommonliftActivities',8,{
                                            Activity : arr18[6].milestone,
                                            TowerA  : arr18[6].count,
                                            TowerB  : arr19[6].count,
                                            TowerC  : arr20[6].count,
                                            TowerD  : arr21[6].count,
                                            TowerE  : arr22[6].count,
                                            TowerF :  arr23[6].count,
                                            TowerG :  arr24[6].count,
                                            TowerH :  arr25[6].count,        
                                            CompletedPercent : ((arr18[6]['count']+arr19[6]['count']+arr20[6]['count']+arr21[6]['count']+arr22[6]['count']+arr23[6]['count']+arr24[6]['count']+arr25[6]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                          });
                                    });
                              });
                        });
                  });
              });
          });
      });
  });
  return;
}


//function to generate excel for property Southren Crest
function commonlobbyExcelSC(floor_count,arr6,arr7,arr8)
{
    /*NOTE:
     *arr6 is arr12, arr7 is arr13 and arr8 is arr14
     */

        //path to generate excel
        const path = require('path');
        let excel = new Excel(path.join(__dirname,'SouthrenCrest.xlsx'));
        
        excel.writeSheet('CommonLobbyActivities',['Activity','TowerA','TowerB','TowerC','CompletedPercent'],[floor_count]).then(()=>{
            excel.writeRow('CommonLobbyActivities',1,{
               Activity: "Unit Count",
               TowerA : floor_count[13]['total_unit'],
               TowerB : floor_count[14]['total_unit'],
               TowerC : floor_count[15]['total_unit'],
           }).then(()=>{//1st then
               excel.writeRow('CommonLobbyActivities',2,{
                Activity: arr6[0].milestone,
                TowerA : arr6[0].count,
                TowerB : arr7[0].count,
                TowerC : arr8[0].count,
                CompletedPercent : (((parseInt(arr6[0]['count'],10)+parseInt(arr7[0]['count'],10)+parseInt(arr8[0]['count'],10))/(floor_count[13]['total_unit']+floor_count[14]['total_unit']+floor_count[15]['total_unit']))*100).toFixed(2)+"%"
            }).then(()=>{//2nd then
                excel.writeRow('CommonLobbyActivities',3,{
                    Activity: arr6[1].milestone,
                    TowerA : arr6[1].count,
                    TowerB : arr7[1].count,
                    TowerC : arr8[1].count,
                    CompletedPercent : (((parseInt(arr6[1]['count'],10)+parseInt(arr7[1]['count'],10)+parseInt(arr8[1]['count'],10))/(floor_count[13]['total_unit']+floor_count[14]['total_unit']+floor_count[15]['total_unit']))*100).toFixed(2)+"%"
                }).then(()=>{//3rd then
                    excel.writeRow('CommonLobbyActivities',4,{
                        Activity: arr6[2].milestone,
                        TowerA : arr6[2].count,
                        TowerB : arr7[2].count,
                        TowerC : arr8[2].count,
                        CompletedPercent : (((parseInt(arr6[2]['count'],10)+parseInt(arr7[2]['count'],10)+parseInt(arr8[2]['count'],10))/(floor_count[13]['total_unit']+floor_count[14]['total_unit']+floor_count[15]['total_unit']))*100).toFixed(2)+"%"
                    }).then(()=>{//4th then
                        excel.writeRow('CommonLobbyActivities',5,{
                            Activity: arr6[3].milestone,
                            TowerA : arr6[3].count,
                            TowerB : arr7[3].count,
                            TowerC : arr8[3].count,
                            CompletedPercent : (((parseInt(arr6[3]['count'],10)+parseInt(arr7[3]['count'],10)+parseInt(arr8[3]['count'],10))/(floor_count[13]['total_unit']+floor_count[14]['total_unit']+floor_count[15]['total_unit']))*100).toFixed(2)+"%"
                         }).then(()=>{//5th then
                            excel.writeRow('CommonLobbyActivities',6,{
                                Activity: arr6[4].milestone,
                                TowerA : arr6[4].count,
                                TowerB : arr7[4].count,
                                TowerC : arr8[4].count,
                                CompletedPercent : (((parseInt(arr6[4]['count'],10)+parseInt(arr7[4]['count'],10)+parseInt(arr8[4]['count'],10))/(floor_count[13]['total_unit']+floor_count[14]['total_unit']+floor_count[15]['total_unit']))*100).toFixed(2)+"%"
                             }).then(()=>{//6th then
                                excel.writeRow('CommonLobbyActivities',7,{
                                    Activity: arr6[5].milestone,
                                    TowerA : arr6[5].count,
                                    TowerB : arr7[5].count,
                                    TowerC : arr8[5].count,
                                    CompletedPercent : (((parseInt(arr6[5]['count'],10)+parseInt(arr7[5]['count'],10)+parseInt(arr8[5]['count'],10))/(floor_count[13]['total_unit']+floor_count[14]['total_unit']+floor_count[15]['total_unit']))*100).toFixed(2)+"%"
                              }).then(()=>{//7th then
                                excel.writeRow('CommonLobbyActivities',7,{
                                    Activity: arr6[6].milestone,
                                    TowerA : arr6[6].count,
                                    TowerB : arr7[6].count,
                                    TowerC : arr8[6].count,
                                    CompletedPercent : (((parseInt(arr6[6]['count'],10)+parseInt(arr7[6]['count'],10)+parseInt(arr8[6]['count'],10))/(floor_count[13]['total_unit']+floor_count[14]['total_unit']+floor_count[15]['total_unit']))*100).toFixed(2)+"%"
                                })
                              })//7th
                            });//6th then ending
                        });//5th then ening
                    });//4th then ending                   
                });//3rd then ending   
             });//2nd then ending
           });//1st then ending
        });
  return;    
}


//function for excel generation of Luxor
function commonlobbyExcelLuxor(floor_count,arr6,arr7,arr8,arr9,arr10,arr11,arr12,arr13)
{
    const path = require('path');
    let excel = new Excel(path.join(__dirname,'Luxor.xlsx'));

    excel.writeSheet('CommonLobbyActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','TowerG','TowerH','CompletedPercent'],[floor_count]).then(()=>{
        excel.writeRow('CommonLobbyActivities',1,{
           Activity: "Floor Count",
           TowerA : floor_count[16]['total_unit'],
           TowerB : floor_count[17]['total_unit'],
           TowerC : floor_count[18]['total_unit'],
           TowerD : floor_count[19]['total_unit'],
           TowerE : floor_count[20]['total_unit'],
           TowerF : floor_count[21]['total_unit'],
           TowerG : floor_count[22]['total_unit'],
           TowerH : floor_count[23]['total_unit'],
       }).then(()=>{//1st then
           excel.writeRow('CommonLobbyActivities',2,{
            Activity: arr6[0].milestone,
            TowerA : arr6[0].count,
            TowerB : arr7[0].count,
            TowerC : arr8[0].count,
            TowerD : arr9[0].count,
            TowerE : arr10[0].count,
            TowerF : arr11[0].count,
            TowerG : arr12[0].count,
            TowerH : arr13[0].count,
            CompletedPercent : (((parseInt(arr6[0]['count'],10)+parseInt(arr7[0]['count'],10)+parseInt(arr8[0]['count'],10)+parseInt(arr9[0]['count'],10)+parseInt(arr10[0]['count'],10)+parseInt(arr11[0]['count'],10)+parseInt(arr12[0]['count'],10)+parseInt(arr13[0]['count'],10))/(floor_count[16]['total_unit']+floor_count[17]['total_unit']+floor_count[18]['total_unit']+floor_count[19]['total_unit']+floor_count[20]['total_unit']+floor_count[21]['total_unit']+floor_count[22]['total_unit']+floor_count[23]['total_unit']))*100).toFixed(2)+"%"
         }).then(()=>{//2nd then
            excel.writeRow('CommonLobbyActivities',3,{
                Activity: arr6[1].milestone,
                TowerA : arr6[1].count,
                TowerB : arr7[1].count,
                TowerC : arr8[1].count,
                TowerD : arr9[1].count,
                TowerE : arr10[1].count,
                TowerF : arr11[1].count,
                TowerG : arr12[1].count,
                TowerH : arr13[1].count,    
                CompletedPercent : (((parseInt(arr6[1]['count'],10)+parseInt(arr7[1]['count'],10)+parseInt(arr8[1]['count'],10)+parseInt(arr9[1]['count'],10)+parseInt(arr10[1]['count'],10)+parseInt(arr11[1]['count'],10)+parseInt(arr12[1]['count'],10)+parseInt(arr13[1]['count'],10))/(floor_count[16]['total_unit']+floor_count[17]['total_unit']+floor_count[18]['total_unit']+floor_count[19]['total_unit']+floor_count[20]['total_unit']+floor_count[21]['total_unit']+floor_count[22]['total_unit']+floor_count[23]['total_unit']))*100).toFixed(2)+"%"
                }).then(()=>{//3rd then
                excel.writeRow('CommonLobbyActivities',4,{
                    Activity: arr6[2].milestone,
                    TowerA : arr6[2].count,
                    TowerB : arr7[2].count,
                    TowerC : arr8[2].count,
                    TowerD : arr9[2].count,
                    TowerE : arr10[2].count,
                    TowerF : arr11[2].count,
                    TowerG : arr12[2].count,
                    TowerH : arr13[2].count,        
                    CompletedPercent : (((parseInt(arr6[2]['count'],10)+parseInt(arr7[2]['count'],10)+parseInt(arr8[2]['count'],10)+parseInt(arr9[2]['count'],10)+parseInt(arr10[2]['count'],10)+parseInt(arr11[2]['count'],10)+parseInt(arr12[2]['count'],10)+parseInt(arr13[2]['count'],10))/(floor_count[16]['total_unit']+floor_count[17]['total_unit']+floor_count[18]['total_unit']+floor_count[19]['total_unit']+floor_count[20]['total_unit']+floor_count[21]['total_unit']+floor_count[22]['total_unit']+floor_count[23]['total_unit']))*100).toFixed(2)+"%"
                    }).then(()=>{//4th then
                    excel.writeRow('CommonLobbyActivities',5,{
                        Activity: arr6[3].milestone,
                        TowerA : arr6[3].count,
                        TowerB : arr7[3].count,
                        TowerC : arr8[3].count,
                        TowerD : arr9[3].count,
                        TowerE : arr10[3].count,
                        TowerF : arr11[3].count,
                        TowerG : arr12[3].count,
                        TowerH : arr13[3].count,
                        CompletedPercent : (((parseInt(arr6[3]['count'],10)+parseInt(arr7[3]['count'],10)+parseInt(arr8[3]['count'],10)+parseInt(arr9[3]['count'],10)+parseInt(arr10[3]['count'],10)+parseInt(arr11[3]['count'],10)+parseInt(arr12[3]['count'],10)+parseInt(arr13[3]['count'],10))/(floor_count[16]['total_unit']+floor_count[17]['total_unit']+floor_count[18]['total_unit']+floor_count[19]['total_unit']+floor_count[20]['total_unit']+floor_count[21]['total_unit']+floor_count[22]['total_unit']+floor_count[23]['total_unit']))*100).toFixed(2)+"%"
                        }).then(()=>{//5th then
                        excel.writeRow('CommonLobbyActivities',6,{
                            Activity: arr6[4].milestone,
                            TowerA : arr6[4].count,
                            TowerB : arr7[4].count,
                            TowerC : arr8[4].count,
                            TowerD : arr9[4].count,
                            TowerE : arr10[4].count,
                            TowerF : arr11[4].count,
                            TowerG : arr12[4].count,
                            TowerH : arr13[4].count,    
                            CompletedPercent : (((parseInt(arr6[4]['count'],10)+parseInt(arr7[4]['count'],10)+parseInt(arr8[4]['count'],10)+parseInt(arr9[4]['count'],10)+parseInt(arr10[4]['count'],10)+parseInt(arr11[4]['count'],10)+parseInt(arr12[4]['count'],10)+parseInt(arr13[4]['count'],10))/(floor_count[16]['total_unit']+floor_count[17]['total_unit']+floor_count[18]['total_unit']+floor_count[19]['total_unit']+floor_count[20]['total_unit']+floor_count[21]['total_unit']+floor_count[22]['total_unit']+floor_count[23]['total_unit']))*100).toFixed(2)+"%"
                            }).then(()=>{//6th then
                            excel.writeRow('CommonLobbyActivities',7,{
                                Activity: arr6[5].milestone,
                                TowerA : arr6[5].count,
                                TowerB : arr7[5].count,
                                TowerC : arr8[5].count,
                                TowerD : arr9[5].count,
                                TowerE : arr10[5].count,
                                TowerF : arr11[5].count,
                                TowerG : arr12[5].count,
                                TowerH : arr13[5].count,        
                                CompletedPercent : (((parseInt(arr6[5]['count'],10)+parseInt(arr7[5]['count'],10)+parseInt(arr8[5]['count'],10)+parseInt(arr9[5]['count'],10)+parseInt(arr10[5]['count'],10)+parseInt(arr11[5]['count'],10)+parseInt(arr12[5]['count'],10)+parseInt(arr13[5]['count'],10))/(floor_count[16]['total_unit']+floor_count[17]['total_unit']+floor_count[18]['total_unit']+floor_count[19]['total_unit']+floor_count[20]['total_unit']+floor_count[21]['total_unit']+floor_count[22]['total_unit']+floor_count[23]['total_unit']))*100).toFixed(2)+"%"
                                })                                
                                });//6th then ending
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;    
}



//function to generate excel for commonlobby activities greenfield phase 1
function commonlobbyExcelGreenfield1(floor_count,arr6,arr7,arr8,arr9,arr10,arr11)
{

        //path to generate excel
        const path = require('path');
        let excel = new Excel(path.join(__dirname,'Greenfieldphase1.xlsx'));
        
        excel.writeSheet('CommonLobbyActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','CompletedPercent'],[floor_count]).then(()=>{
            excel.writeRow('CommonLobbyActivities',1,{
               Activity: "Unit Count",
               TowerA : floor_count[4]['total_unit'],
               TowerB : floor_count[5]['total_unit'],
               TowerC : floor_count[6]['total_unit'],
               TowerD : floor_count[7]['total_unit'],
               TowerE : floor_count[8]['total_unit'],
               TowerF : floor_count[9]['total_unit'],               
           }).then(()=>{//1st then
               excel.writeRow('CommonLobbyActivities',2,{
                Activity: arr6[0].milestone,
                TowerA : arr6[0].count,
                TowerB : arr7[0].count,
                TowerC : arr8[0].count,
                TowerD : arr9[0].count,
                TowerE : arr10[0].count,
                TowerF : arr11[0].count,
                CompletedPercent : (((parseInt(arr6[0]['count'],10)+parseInt(arr7[0]['count'],10)+parseInt(arr8[0]['count'],10)+parseInt(arr9[0]['count'],10)+parseInt(arr10[0]['count'],10)+parseInt(arr11[0]['count'],10))/(floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']+floor_count[8]['total_unit']+floor_count[9]['total_unit']))*100).toFixed(2)+"%"
            }).then(()=>{//2nd then
                excel.writeRow('CommonLobbyActivities',3,{
                    Activity: arr6[1].milestone,
                    TowerA : arr6[1].count,
                    TowerB : arr7[1].count,
                    TowerC : arr8[1].count,
                    TowerD : arr9[1].count,
                    TowerE : arr10[1].count,
                    TowerF : arr11[1].count,
                    CompletedPercent : (((parseInt(arr6[1]['count'],10)+parseInt(arr7[1]['count'],10)+parseInt(arr8[1]['count'],10)+parseInt(arr9[1]['count'],10)+parseInt(arr10[1]['count'],10)+parseInt(arr11[1]['count'],10))/(floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']+floor_count[8]['total_unit']+floor_count[9]['total_unit']))*100).toFixed(2)+"%"
                    }).then(()=>{//3rd then
                    excel.writeRow('CommonLobbyActivities',4,{
                        Activity: arr6[2].milestone,
                        TowerA : arr6[2].count,
                        TowerB : arr7[2].count,
                        TowerC : arr8[2].count,
                        TowerD : arr9[2].count,
                        TowerE : arr10[2].count,
                        TowerF : arr11[2].count,
                        CompletedPercent : (((parseInt(arr6[2]['count'],10)+parseInt(arr7[2]['count'],10)+parseInt(arr8[2]['count'],10)+parseInt(arr9[2]['count'],10)+parseInt(arr10[2]['count'],10)+parseInt(arr11[2]['count'],10))/(floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']+floor_count[8]['total_unit']+floor_count[9]['total_unit']))*100).toFixed(2)+"%"
                            }).then(()=>{//4th then
                        excel.writeRow('CommonLobbyActivities',5,{
                            Activity: arr6[3].milestone,
                            TowerA : arr6[3].count,
                            TowerB : arr7[3].count,
                            TowerC : arr8[3].count,
                            TowerD : arr9[3].count,
                            TowerE : arr10[3].count,
                            TowerF : arr11[3].count,
                            CompletedPercent : (((parseInt(arr6[3]['count'],10)+parseInt(arr7[3]['count'],10)+parseInt(arr8[3]['count'],10)+parseInt(arr9[3]['count'],10)+parseInt(arr10[3]['count'],10)+parseInt(arr11[3]['count'],10))/(floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']+floor_count[8]['total_unit']+floor_count[9]['total_unit']))*100).toFixed(2)+"%"
                            }).then(()=>{//5th then
                            excel.writeRow('CommonLobbyActivities',6,{
                                Activity: arr6[4].milestone,
                                TowerA : arr6[4].count,
                                TowerB : arr7[4].count,
                                TowerC : arr8[4].count,
                                TowerD : arr9[4].count,
                                TowerE : arr10[4].count,
                                TowerF : arr11[4].count,
                                CompletedPercent : (((parseInt(arr6[4]['count'],10)+parseInt(arr7[4]['count'],10)+parseInt(arr8[4]['count'],10)+parseInt(arr9[4]['count'],10)+parseInt(arr10[4]['count'],10)+parseInt(arr11[4]['count'],10))/(floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']+floor_count[8]['total_unit']+floor_count[9]['total_unit']))*100).toFixed(2)+"%"
                                }).then(()=>{//6th then
                                excel.writeRow('CommonLobbyActivities',7,{
                                    Activity: arr6[5].milestone,
                                    TowerA : arr6[5].count,
                                    TowerB : arr7[5].count,
                                    TowerC : arr8[5].count,
                                    TowerD : arr9[5].count,
                                    TowerE : arr10[5].count,
                                    TowerF : arr11[5].count,
                                    CompletedPercent : (((parseInt(arr6[5]['count'],10)+parseInt(arr7[5]['count'],10)+parseInt(arr8[5]['count'],10)+parseInt(arr9[5]['count'],10)+parseInt(arr10[5]['count'],10)+parseInt(arr11[5]['count'],10))/(floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']+floor_count[8]['total_unit']+floor_count[9]['total_unit']))*100).toFixed(2)+"%"
                              }).then(() => {//7th then
                                    excel.writeRow('CommonLobbyActivities',8,{
                                    Activity: arr6[6].milestone,
                                    TowerA : arr6[6].count,
                                    TowerB : arr7[6].count,
                                    TowerC : arr8[6].count,
                                    TowerD : arr9[6].count,
                                    TowerE : arr10[6].count,
                                    TowerF : arr11[6].count,
                                    CompletedPercent : (((parseInt(arr6[6]['count'],10)+parseInt(arr7[6]['count'],10)+parseInt(arr8[6]['count'],10)+parseInt(arr9[6]['count'],10)+parseInt(arr10[6]['count'],10)+parseInt(arr11[6]['count'],10))/(floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']+floor_count[8]['total_unit']+floor_count[9]['total_unit']))*100).toFixed(2)+"%"
                                })
                              });//7th then ending
                            });//6th then ending
                        });//5th then ening
                    });//4th then ending                   
                });//3rd then ending   
             });//2nd then ending
           });//1st then ending
        });
  return;    
}


//function to generate excel for commonlobby activities greenfield phase 2
function commonlobbyExcelGreenfield2(floor_count,arr6,arr7,arr8)
{
    /*NOTE:
     *arr6 is arr12, arr7 is arr13 and arr8 is arr14
     */

        //path to generate excel
        const path = require('path');
        let excel = new Excel(path.join(__dirname,'Greenfieldphase2.xlsx'));
        
        excel.writeSheet('CommonLobbyActivities',['Activity','TowerG','TowerH','TowerJ','CompletedPercent'],[floor_count]).then(()=>{
            excel.writeRow('CommonLobbyActivities',1,{
               Activity: "Unit Count",
               TowerG : floor_count[10]['total_unit'],
               TowerH : floor_count[11]['total_unit'],
               TowerJ : floor_count[12]['total_unit'],
           }).then(()=>{//1st then
               excel.writeRow('CommonLobbyActivities',2,{
                Activity: arr6[0].milestone,
                TowerG : arr6[0].count,
                TowerH : arr7[0].count,
                TowerJ : arr8[0].count,
                CompletedPercent : (((parseInt(arr6[0]['count'],10)+parseInt(arr7[0]['count'],10)+parseInt(arr8[0]['count'],10))/(floor_count[10]['total_unit']+floor_count[11]['total_unit']+floor_count[12]['total_unit']))*100).toFixed(2)+"%"
            }).then(()=>{//2nd then
                excel.writeRow('CommonLobbyActivities',3,{
                    Activity: arr6[1].milestone,
                    TowerG : arr6[1].count,
                    TowerH : arr7[1].count,
                    TowerJ : arr8[1].count,
                    CompletedPercent : (((parseInt(arr6[1]['count'],10)+parseInt(arr7[1]['count'],10)+parseInt(arr8[1]['count'],10))/(floor_count[10]['total_unit']+floor_count[11]['total_unit']+floor_count[12]['total_unit']))*100).toFixed(2)+"%"
                }).then(()=>{//3rd then
                    excel.writeRow('CommonLobbyActivities',4,{
                        Activity: arr6[2].milestone,
                        TowerG : arr6[2].count,
                        TowerH : arr7[2].count,
                        TowerJ : arr8[2].count,
                        CompletedPercent : (((parseInt(arr6[2]['count'],10)+parseInt(arr7[2]['count'],10)+parseInt(arr8[2]['count'],10))/(floor_count[10]['total_unit']+floor_count[11]['total_unit']+floor_count[12]['total_unit']))*100).toFixed(2)+"%"
                    }).then(()=>{//4th then
                        excel.writeRow('CommonLobbyActivities',5,{
                            Activity: arr6[3].milestone,
                            TowerG : arr6[3].count,
                            TowerH : arr7[3].count,
                            TowerJ : arr8[3].count,
                            CompletedPercent : (((parseInt(arr6[3]['count'],10)+parseInt(arr7[3]['count'],10)+parseInt(arr8[3]['count'],10))/(floor_count[10]['total_unit']+floor_count[11]['total_unit']+floor_count[12]['total_unit']))*100).toFixed(2)+"%"
                         }).then(()=>{//5th then
                            excel.writeRow('CommonLobbyActivities',6,{
                                Activity: arr6[4].milestone,
                                TowerG : arr6[4].count,
                                TowerH : arr7[4].count,
                                TowerJ : arr8[4].count,
                                CompletedPercent : (((parseInt(arr6[4]['count'],10)+parseInt(arr7[4]['count'],10)+parseInt(arr8[4]['count'],10))/(floor_count[10]['total_unit']+floor_count[11]['total_unit']+floor_count[12]['total_unit']))*100).toFixed(2)+"%"
                             }).then(()=>{//6th then
                                excel.writeRow('CommonLobbyActivities',7,{
                                    Activity: arr6[5].milestone,
                                    TowerG : arr6[5].count,
                                    TowerH : arr7[5].count,
                                    TowerJ : arr8[5].count,
                                    CompletedPercent : (((parseInt(arr6[5]['count'],10)+parseInt(arr7[5]['count'],10)+parseInt(arr8[5]['count'],10))/(floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']+floor_count[8]['total_unit']+floor_count[9]['total_unit']))*100).toFixed(2)+"%"
                              })
                            });//6th then ending
                        });//5th then ening
                    });//4th then ending                   
                });//3rd then ending   
             });//2nd then ending
           });//1st then ending
        });
  return;    
}


//function to generate excel for flatfinish Activities vizag
function flatfinishExcelVizag(floor_count,arr1,arr2,arr3,arr4)
{
        //path to generate excel
        const path = require('path');
        let excel = new Excel(path.join(__dirname,'PanoramaHills.xlsx'));

        excel.writeSheet('FlatFinishActivities',['Activity','Tower2A','Tower2B','Tower2C','Tower2D','CompletedPercent'],[floor_count]).then(()=>{
            excel.writeRow('FlatFinishActivities',1,{
               Activity: "Floor Count",
               Tower2A : floor_count[0]['total_unit'],
               Tower2B : floor_count[1]['total_unit'],
               Tower2C : floor_count[2]['total_unit'],
               Tower2D : floor_count[3]['total_unit'],
               
           }).then(()=>{//1st then
               excel.writeRow('FlatFinishActivities',2,{
                Activity: arr1[0].milestone,
                Tower2A : arr1[0].count,
                Tower2B : arr2[0].count,
                Tower2C : arr3[0].count,
                Tower2D : arr4[0].count,
                CompletedPercent : (((parseInt(arr1[0]['count'],10)+parseInt(arr2[0]['count'],10)+parseInt(arr3[0]['count'],10)+parseInt(arr4[0]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"
             }).then(()=>{//2nd then
                excel.writeRow('FlatFinishActivities',3,{
                    Activity: arr1[1].milestone,
                    Tower2A : arr1[1].count,
                    Tower2B : arr2[1].count,
                    Tower2C : arr3[1].count,
                    Tower2D : arr4[1].count,
                    CompletedPercent : (((parseInt(arr1[1]['count'],10)+parseInt(arr2[1]['count'],10)+parseInt(arr3[1]['count'],10)+parseInt(arr4[1]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"
                }).then(()=>{//3rd then
                    excel.writeRow('FlatFinishActivities',4,{
                        Activity: arr1[2].milestone,
                        Tower2A : arr1[2].count,
                        Tower2B : arr2[2].count,
                        Tower2C : arr3[2].count,
                        Tower2D : arr4[2].count,
                        CompletedPercent : (((parseInt(arr1[2]['count'],10)+parseInt(arr2[2]['count'],10)+parseInt(arr3[2]['count'],10)+parseInt(arr4[2]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"
    
                    }).then(()=>{//4th then
                        excel.writeRow('FlatFinishActivities',5,{
                            Activity: arr1[3].milestone,
                            Tower2A : arr1[3].count,
                            Tower2B : arr2[3].count,
                            Tower2C : arr3[3].count,
                            Tower2D : arr4[3].count,
                            CompletedPercent : (((parseInt(arr1[3]['count'],10)+parseInt(arr2[3]['count'],10)+parseInt(arr3[3]['count'],10)+parseInt(arr4[3]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"
                                    
                        }).then(()=>{//5th then
                            excel.writeRow('FlatFinishActivities',6,{
                                Activity: arr1[4].milestone,
                                Tower2A : arr1[4].count,
                                Tower2B : arr2[4].count,
                                Tower2C : arr3[4].count,
                                Tower2D : arr4[4].count,  
                                CompletedPercent : (((parseInt(arr1[4]['count'],10)+parseInt(arr2[4]['count'],10)+parseInt(arr3[4]['count'],10)+parseInt(arr4[4]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"  
        
                            }).then(()=>{//6th then
                                excel.writeRow('FlatFinishActivities',7,{
                                    Activity: arr1[5].milestone,
                                    Tower2A : arr1[5].count,
                                    Tower2B : arr2[5].count,
                                    Tower2C : arr3[5].count,
                                    Tower2D : arr4[5].count,
                                    CompletedPercent : (((parseInt(arr1[5]['count'],10)+parseInt(arr2[5]['count'],10)+parseInt(arr3[5]['count'],10)+parseInt(arr4[5]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                }).then(()=>{//7th then     
                                    excel.writeRow('FlatFinishActivities',8,{
                                        Activity: arr1[6].milestone,
                                        Tower2A : arr1[6].count,
                                        Tower2B : arr2[6].count,
                                        Tower2C : arr3[6].count,
                                        Tower2D : arr4[6].count,
                                        CompletedPercent : (((parseInt(arr1[6]['count'],10)+parseInt(arr2[6]['count'],10)+parseInt(arr3[6]['count'],10)+parseInt(arr4[6]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                    }).then(()=>{//8th then
                                        excel.writeRow('FlatFinishActivities',9,{
                                            Activity: arr1[7].milestone,
                                            Tower2A : arr1[7].count,
                                            Tower2B : arr2[7].count,
                                            Tower2C : arr3[7].count,
                                            Tower2D : arr4[7].count,
                                            CompletedPercent : (((parseInt(arr1[7]['count'],10)+parseInt(arr2[7]['count'],10)+parseInt(arr3[7]['count'],10)+parseInt(arr4[7]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                        }).then(()=>{//9th then
                                            excel.writeRow('FlatFinishActivities',10,{
                                                Activity: arr1[8].milestone,
                                                Tower2A : arr1[8].count,
                                                Tower2B : arr2[8].count,
                                                Tower2C : arr3[8].count,
                                                Tower2D : arr4[8].count,
                                                CompletedPercent : (((parseInt(arr1[8]['count'],10)+parseInt(arr2[8]['count'],10)+parseInt(arr3[8]['count'],10)+parseInt(arr4[8]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                            }).then(()=>{//10th then
                                                excel.writeRow('FlatFinishActivities',11,{
                                                    Activity: arr1[9].milestone,
                                                    Tower2A : arr1[9].count,
                                                    Tower2B : arr2[9].count,
                                                    Tower2C : arr3[9].count,
                                                    Tower2D : arr4[9].count,
                                                    CompletedPercent : (((parseInt(arr1[9]['count'],10)+parseInt(arr2[9]['count'],10)+parseInt(arr3[9]['count'],10)+parseInt(arr4[9]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                }).then(()=>{//11th then
                                                    excel.writeRow('FlatFinishActivities',12,{
                                                        Activity: arr1[10].milestone,
                                                        Tower2A : arr1[10].count,
                                                        Tower2B : arr2[10].count,
                                                        Tower2C : arr3[10].count,
                                                        Tower2D : arr4[10].count,
                                                        CompletedPercent : (((parseInt(arr1[10]['count'],10)+parseInt(arr2[10]['count'],10)+parseInt(arr3[10]['count'],10)+parseInt(arr4[10]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                    }).then(()=>{//12th then
                                                        excel.writeRow('FlatFinishActivities',13,{
                                                            Activity: arr1[11].milestone,
                                                            Tower2A : arr1[11].count,
                                                            Tower2B : arr2[11].count,
                                                            Tower2C : arr3[11].count,
                                                            Tower2D : arr4[11].count,
                                                            CompletedPercent : (((parseInt(arr1[11]['count'],10)+parseInt(arr2[11]['count'],10)+parseInt(arr3[11]['count'],10)+parseInt(arr4[11]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                        }).then(()=>{//13th then
                                                            excel.writeRow('FlatFinishActivities',14,{
                                                                Activity: arr1[12].milestone,
                                                                Tower2A : arr1[12].count,
                                                                Tower2B : arr2[12].count,
                                                                Tower2C : arr3[12].count,
                                                                Tower2D : arr4[12].count,
                                                                CompletedPercent : (((parseInt(arr1[12]['count'],10)+parseInt(arr2[12]['count'],10)+parseInt(arr3[12]['count'],10)+parseInt(arr4[12]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                            }).then(()=>{//14th then
                                                                excel.writeRow('FlatFinishActivities',15,{
                                                                    Activity: arr1[13].milestone,
                                                                    Tower2A : arr1[13].count,
                                                                    Tower2B : arr2[13].count,
                                                                    Tower2C : arr3[13].count,
                                                                    Tower2D : arr4[13].count,
                                                                    CompletedPercent : (((parseInt(arr1[13]['count'],10)+parseInt(arr2[13]['count'],10)+parseInt(arr3[13]['count'],10)+parseInt(arr4[13]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                }).then(()=>{//15th then
                                                                    excel.writeRow('FlatFinishActivities',16,{
                                                                        Activity: arr1[14].milestone,
                                                                        Tower2A : arr1[14].count,
                                                                        Tower2B : arr2[14].count,
                                                                        Tower2C : arr3[14].count,
                                                                        Tower2D : arr4[14].count,
                                                                        CompletedPercent : (((parseInt(arr1[14]['count'],10)+parseInt(arr2[14]['count'],10)+parseInt(arr3[14]['count'],10)+parseInt(arr4[14]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                    }).then(()=>{//16th
                                                                        excel.writeRow('FlatFinishActivities',17,{
                                                                            Activity: arr1[15].milestone,
                                                                            Tower2A : arr1[15].count,
                                                                            Tower2B : arr2[15].count,
                                                                            Tower2C : arr3[15].count,
                                                                            Tower2D : arr4[15].count,
                                                                            CompletedPercent : (((parseInt(arr1[15]['count'],10)+parseInt(arr2[15]['count'],10)+parseInt(arr3[15]['count'],10)+parseInt(arr4[15]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                        }).then(()=>{//17th
                                                                            excel.writeRow('FlatFinishActivities',18,{
                                                                                Activity: arr1[16].milestone,
                                                                                Tower2A : arr1[16].count,
                                                                                Tower2B : arr2[16].count,
                                                                                Tower2C : arr3[16].count,
                                                                                Tower2D : arr4[16].count,
                                                                                CompletedPercent : (((parseInt(arr1[16]['count'],10)+parseInt(arr2[16]['count'],10)+parseInt(arr3[16]['count'],10)+parseInt(arr4[16]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                            }).then(()=>{//18th
                                                                                excel.writeRow('FlatFinishActivities',19,{
                                                                                    Activity: arr1[17].milestone,
                                                                                    Tower2A : arr1[17].count,
                                                                                    Tower2B : arr2[17].count,
                                                                                    Tower2C : arr3[17].count,
                                                                                    Tower2D : arr4[17].count,
                                                                                    CompletedPercent : (((parseInt(arr1[17]['count'],10)+parseInt(arr2[17]['count'],10)+parseInt(arr3[17]['count'],10)+parseInt(arr4[17]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                                }).then(()=>{//19th
                                                                                    excel.writeRow('FlatFinishActivities',20,{
                                                                                        Activity: arr1[18].milestone,
                                                                                        Tower2A : arr1[18].count,
                                                                                        Tower2B : arr2[18].count,
                                                                                        Tower2C : arr3[18].count,
                                                                                        Tower2D : arr4[18].count,
                                                                                        CompletedPercent : (((parseInt(arr1[18]['count'],10)+parseInt(arr2[18]['count'],10)+parseInt(arr3[18]['count'],10)+parseInt(arr4[18]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                                    }).then(()=>{//20th
                                                                                        excel.writeRow('FlatFinishActivities',21,{
                                                                                            Activity: arr1[19].milestone,
                                                                                            Tower2A : arr1[19].count,
                                                                                            Tower2B : arr2[19].count,
                                                                                            Tower2C : arr3[19].count,
                                                                                            Tower2D : arr4[19].count,
                                                                                            CompletedPercent : (((parseInt(arr1[19]['count'],10)+parseInt(arr2[19]['count'],10)+parseInt(arr3[19]['count'],10)+parseInt(arr4[19]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                                        }).then(() => {//21st
                                                                                            excel.writeRow('FlatFinishActivities',22,{
                                                                                                Activity: arr1[20].milestone,
                                                                                                Tower2A : arr1[20].count,
                                                                                                Tower2B : arr2[20].count,
                                                                                                Tower2C : arr3[20].count,
                                                                                                Tower2D : arr4[20].count,
                                                                                                CompletedPercent : (((parseInt(arr1[20]['count'],10)+parseInt(arr2[20]['count'],10)+parseInt(arr3[20]['count'],10)+parseInt(arr4[20]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                                            }).then(() => {//22nd
                                                                                                excel.writeRow('FlatFinishActivities',23,{
                                                                                                    Activity: arr1[21].milestone,
                                                                                                    Tower2A : arr1[21].count,
                                                                                                    Tower2B : arr2[21].count,
                                                                                                    Tower2C : arr3[21].count,
                                                                                                    Tower2D : arr4[21].count,
                                                                                                    CompletedPercent : (((parseInt(arr1[21]['count'],10)+parseInt(arr2[21]['count'],10)+parseInt(arr3[21]['count'],10)+parseInt(arr4[21]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                                                }).then(() => {//23rd
                                                                                                    excel.writeRow('FlatFinishActivities',24,{
                                                                                                        Activity: arr1[22].milestone,
                                                                                                        Tower2A : arr1[22].count,
                                                                                                        Tower2B : arr2[22].count,
                                                                                                        Tower2C : arr3[22].count,
                                                                                                        Tower2D : arr4[22].count,
                                                                                                        CompletedPercent : (((parseInt(arr1[22]['count'],10)+parseInt(arr2[22]['count'],10)+parseInt(arr3[22]['count'],10)+parseInt(arr4[22]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                                                    }).then(() => {//24th
                                                                                                        excel.writeRow('FlatFinishActivities',25,{
                                                                                                            Activity: arr1[23].milestone,
                                                                                                            Tower2A : arr1[23].count,
                                                                                                            Tower2B : arr2[23].count,
                                                                                                            Tower2C : arr3[23].count,
                                                                                                            Tower2D : arr4[23].count,
                                                                                                            CompletedPercent : (((parseInt(arr1[23]['count'],10)+parseInt(arr2[23]['count'],10)+parseInt(arr3[23]['count'],10)+parseInt(arr4[23]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                                                        }).then(() => {//25th
                                                                                                            excel.writeRow('FlatFinishActivities',26,{
                                                                                                                Activity: arr1[24].milestone,
                                                                                                                Tower2A : arr1[24].count,
                                                                                                                Tower2B : arr2[24].count,
                                                                                                                Tower2C : arr3[24].count,
                                                                                                                Tower2D : arr4[24].count,
                                                                                                                CompletedPercent : (((parseInt(arr1[24]['count'],10)+parseInt(arr2[24]['count'],10)+parseInt(arr3[24]['count'],10)+parseInt(arr4[24]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                                                             }).then(() =>{//26th
                                                                                                                excel.writeRow('FlatFinishActivities',27,{
                                                                                                                    Activity: arr1[25].milestone,
                                                                                                                    Tower2A : arr1[25].count,
                                                                                                                    Tower2B : arr2[25].count,
                                                                                                                    Tower2C : arr3[25].count,
                                                                                                                    Tower2D : arr4[25].count,
                                                                                                                    CompletedPercent : (((parseInt(arr1[25]['count'],10)+parseInt(arr2[25]['count'],10)+parseInt(arr3[25]['count'],10)+parseInt(arr4[25]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"           
                                                                                                                 }).then(()=>{//27th
                                                                                                                    excel.writeRow('FlatFinishActivities',28,{
                                                                                                                        Activity: arr1[26].milestone,
                                                                                                                        Tower2A : arr1[26].count,
                                                                                                                        Tower2B : arr2[26].count,
                                                                                                                        Tower2C : arr3[26].count,
                                                                                                                        Tower2D : arr4[26].count,
                                                                                                                        CompletedPercent : (((parseInt(arr1[26]['count'],10)+parseInt(arr2[26]['count'],10)+parseInt(arr3[26]['count'],10)+parseInt(arr4[26]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                }).then(() =>{//28th
                                                                                                                    excel.writeRow('FlatFinishActivities',29,{
                                                                                                                        Activity: arr1[27].milestone,
                                                                                                                        Tower2A : arr1[27].count,
                                                                                                                        Tower2B : arr2[27].count,
                                                                                                                        Tower2C : arr3[27].count,
                                                                                                                        Tower2D : arr4[27].count,
                                                                                                                        CompletedPercent : (((parseInt(arr1[27]['count'],10)+parseInt(arr2[27]['count'],10)+parseInt(arr3[27]['count'],10)+parseInt(arr4[27]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                              }).then(()=>{//29th
                                                                                                                excel.writeRow('FlatFinishActivities',30,{
                                                                                                                    Activity: arr1[28].milestone,
                                                                                                                    Tower2A : arr1[28].count,
                                                                                                                    Tower2B : arr2[28].count,
                                                                                                                    Tower2C : arr3[28].count,
                                                                                                                    Tower2D : arr4[28].count,
                                                                                                                    CompletedPercent : (((parseInt(arr1[28]['count'],10)+parseInt(arr2[28]['count'],10)+parseInt(arr3[28]['count'],10)+parseInt(arr4[28]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                }).then(()=>{//30th
                                                                                                                    excel.writeRow('FlatFinishActivities',31,{
                                                                                                                        Activity: arr1[29].milestone,
                                                                                                                        Tower2A : arr1[29].count,
                                                                                                                        Tower2B : arr2[29].count,
                                                                                                                        Tower2C : arr3[29].count,
                                                                                                                        Tower2D : arr4[29].count,
                                                                                                                        CompletedPercent : (((parseInt(arr1[29]['count'],10)+parseInt(arr2[29]['count'],10)+parseInt(arr3[29]['count'],10)+parseInt(arr4[29]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                    }).then(()=>{//31st
                                                                                                                        excel.writeRow('FlatFinishActivities',32,{
                                                                                                                            Activity: arr1[30].milestone,
                                                                                                                            Tower2A : arr1[30].count,
                                                                                                                            Tower2B : arr2[30].count,
                                                                                                                            Tower2C : arr3[30].count,
                                                                                                                            Tower2D : arr4[30].count,
                                                                                                                            CompletedPercent : (((parseInt(arr1[30]['count'],10)+parseInt(arr2[30]['count'],10)+parseInt(arr3[30]['count'],10)+parseInt(arr4[30]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                        }).then(() =>{//32nd
                                                                                                                            excel.writeRow('FlatFinishActivities',33,{
                                                                                                                                Activity: arr1[31].milestone,
                                                                                                                                Tower2A : arr1[31].count,
                                                                                                                                Tower2B : arr2[31].count,
                                                                                                                                Tower2C : arr3[31].count,
                                                                                                                                Tower2D : arr4[31].count,
                                                                                                                                CompletedPercent : (((parseInt(arr1[31]['count'],10)+parseInt(arr2[31]['count'],10)+parseInt(arr3[31]['count'],10)+parseInt(arr4[31]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                            }).then(()=>{//33rd
                                                                                                                                excel.writeRow('FlatFinishActivities',34,{
                                                                                                                                    Activity: arr1[32].milestone,
                                                                                                                                    Tower2A : arr1[32].count,
                                                                                                                                    Tower2B : arr2[32].count,
                                                                                                                                    Tower2C : arr3[32].count,
                                                                                                                                    Tower2D : arr4[32].count,
                                                                                                                                    CompletedPercent : (((parseInt(arr1[32]['count'],10)+parseInt(arr2[32]['count'],10)+parseInt(arr3[32]['count'],10)+parseInt(arr4[32]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                                }).then(() =>{//34th
                                                                                                                                    excel.writeRow('FlatFinishActivities',35,{
                                                                                                                                        Activity: arr1[33].milestone,
                                                                                                                                        Tower2A : arr1[33].count,
                                                                                                                                        Tower2B : arr2[33].count,
                                                                                                                                        Tower2C : arr3[33].count,
                                                                                                                                        Tower2D : arr4[33].count,
                                                                                                                                        CompletedPercent : (((parseInt(arr1[33]['count'],10)+parseInt(arr2[33]['count'],10)+parseInt(arr3[33]['count'],10)+parseInt(arr4[33]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                                    }).then(() => {//35th
                                                                                                                                        excel.writeRow('FlatFinishActivities',36,{
                                                                                                                                            Activity: arr1[34].milestone,
                                                                                                                                            Tower2A : arr1[34].count,
                                                                                                                                            Tower2B : arr2[34].count,
                                                                                                                                            Tower2C : arr3[34].count,
                                                                                                                                            Tower2D : arr4[34].count,
                                                                                                                                            CompletedPercent : (((parseInt(arr1[34]['count'],10)+parseInt(arr2[34]['count'],10)+parseInt(arr3[34]['count'],10)+parseInt(arr4[34]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                                        }).then(() => {//36th
                                                                                                                                            excel.writeRow('FlatFinishActivities',37,{
                                                                                                                                                Activity: arr1[35].milestone,
                                                                                                                                                Tower2A : arr1[35].count,
                                                                                                                                                Tower2B : arr2[35].count,
                                                                                                                                                Tower2C : arr3[35].count,
                                                                                                                                                Tower2D : arr4[35].count,
                                                                                                                                                CompletedPercent : (((parseInt(arr1[35]['count'],10)+parseInt(arr2[35]['count'],10)+parseInt(arr3[35]['count'],10)+parseInt(arr4[35]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                                            }).then(() => {//37th
                                                                                                                                                excel.writeRow('FlatFinishActivities',38,{
                                                                                                                                                    Activity: arr1[36].milestone,
                                                                                                                                                    Tower2A : arr1[36].count,
                                                                                                                                                    Tower2B : arr2[36].count,
                                                                                                                                                    Tower2C : arr3[36].count,
                                                                                                                                                    Tower2D : arr4[36].count,
                                                                                                                                                    CompletedPercent : (((parseInt(arr1[36]['count'],10)+parseInt(arr2[36]['count'],10)+parseInt(arr3[36]['count'],10)+parseInt(arr4[36]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                                                }).then(() => {//38th
                                                                                                                                                    excel.writeRow('FlatFinishActivities',39,{
                                                                                                                                                        Activity: arr1[37].milestone,
                                                                                                                                                        Tower2A : arr1[37].count,
                                                                                                                                                        Tower2B : arr2[37].count,
                                                                                                                                                        Tower2C : arr3[37].count,
                                                                                                                                                        Tower2D : arr4[37].count,
                                                                                                                                                        CompletedPercent : (((parseInt(arr1[37]['count'],10)+parseInt(arr2[37]['count'],10)+parseInt(arr3[37]['count'],10)+parseInt(arr4[37]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']))*100).toFixed(2)+"%"                                                                                                                                
                                                                                                                                                    })
                                                                                                                                                })//38th
                                                                                                                                            })//37th
                                                                                                                                        })//36th
                                                                                                                                    })//35th    
                                                                                                                                })//34th
                                                                                                                            })//33rd
                                                                                                                        })//32nd
                                                                                                                    })//31st
                                                                                                                })//30th
                                                                                                              })//29th
                                                                                                            })//28th
                                                                                                          })//27th
                                                                                                        })//26th
                                                                                                      })//25th
                                                                                                    })//24th
                                                                                                })//23rd 
                                                                                            })//22nd
                                                                                        })//21st
                                                                                    })//20th
                                                                                })//19th then
                                                                            })//18th then 
                                                                        })//17th then
                                                                    })//16th then
                                                                })//15th then ending
                                                            })//14th then ending
                                                        })//13th then ending
                                                    })//12th then ending
                                                })//11th then ending
                                            })//10th then ending
                                        })//9th then ending
                                    })//8th then ending
                                })//7th then ending
                            });//6th then ending
                        });//5th then ening
                    });//4th then ending                   
                });//3rd then ending   
             });//2nd then ending
           });//1st then ending
        });
  return;    

}


//flat finish excel fror phase 1
function  flatfinishExcelGreenfield1(floor_count,arr6,arr7,arr8,arr9,arr10,arr11)
{

            //path to generate excel
            const path = require('path');
            let excel = new Excel(path.join(__dirname,'Greenfieldphase1.xlsx'));
    
            excel.writeSheet('FlatFinishActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','CompletedPercent'],[floor_count]).then(()=>{
                excel.writeRow('FlatFinishActivities',1,{
                   Activity: "Floor Count",
                   TowerA : floor_count[0]['total_unit'],
                   TowerB : floor_count[1]['total_unit'],
                   TowerC : floor_count[2]['total_unit'],
                   TowerD : floor_count[3]['total_unit'],
                   TowerE : floor_count[4]['total_unit'],
                   TowerF : floor_count[5]['total_unit'],
               }).then(()=>{//1st then
                   excel.writeRow('FlatFinishActivities',2,{
                    Activity: arr6[0].milestone,
                    TowerA : arr6[0].count,
                    TowerB : arr7[0].count,
                    TowerC : arr8[0].count,
                    TowerD : arr9[0].count,
                    TowerE : arr10[0].count,
                    TowerF : arr11[0].count,
                    CompletedPercent : (((parseInt(arr6[0]['count'],10)+parseInt(arr7[0]['count'],10)+parseInt(arr8[0]['count'],10)+parseInt(arr9[0]['count'],10)+parseInt(arr10[0]['count'],10)+parseInt(arr11[0]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                 }).then(()=>{//2nd then
                    excel.writeRow('FlatFinishActivities',3,{
                        Activity: arr6[1].milestone,
                        TowerA : arr6[1].count,
                        TowerB : arr7[1].count,
                        TowerC : arr8[1].count,
                        TowerD : arr9[1].count,
                        TowerE : arr10[1].count,
                        TowerF : arr11[1].count,
                        CompletedPercent : (((parseInt(arr6[1]['count'],10)+parseInt(arr7[1]['count'],10)+parseInt(arr8[1]['count'],10)+parseInt(arr9[1]['count'],10)+parseInt(arr10[1]['count'],10)+parseInt(arr11[1]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                        }).then(()=>{//3rd then
                        excel.writeRow('FlatFinishActivities',4,{
                            Activity: arr6[2].milestone,
                            TowerA : arr6[2].count,
                            TowerB : arr7[2].count,
                            TowerC : arr8[2].count,
                            TowerD : arr9[2].count,
                            TowerE : arr10[2].count,
                            TowerF : arr11[2].count,
                            CompletedPercent : (((parseInt(arr6[2]['count'],10)+parseInt(arr7[2]['count'],10)+parseInt(arr8[2]['count'],10)+parseInt(arr9[2]['count'],10)+parseInt(arr10[2]['count'],10)+parseInt(arr11[2]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                            }).then(()=>{//4th then
                            excel.writeRow('FlatFinishActivities',5,{
                                Activity: arr6[3].milestone,
                                TowerA : arr6[3].count,
                                TowerB : arr7[3].count,
                                TowerC : arr8[3].count,
                                TowerD : arr9[3].count,
                                TowerE : arr10[3].count,
                                TowerF : arr11[3].count,
                                CompletedPercent : (((parseInt(arr6[3]['count'],10)+parseInt(arr7[3]['count'],10)+parseInt(arr8[3]['count'],10)+parseInt(arr9[3]['count'],10)+parseInt(arr10[3]['count'],10)+parseInt(arr11[3]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                }).then(()=>{//5th then
                                excel.writeRow('FlatFinishActivities',6,{
                                    Activity: arr6[4].milestone,
                                    TowerA : arr6[4].count,
                                    TowerB : arr7[4].count,
                                    TowerC : arr8[4].count,
                                    TowerD : arr9[4].count,
                                    TowerE : arr10[4].count,
                                    TowerF : arr11[4].count,
                                    CompletedPercent : (((parseInt(arr6[4]['count'],10)+parseInt(arr7[4]['count'],10)+parseInt(arr8[4]['count'],10)+parseInt(arr9[4]['count'],10)+parseInt(arr10[4]['count'],10)+parseInt(arr11[4]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                    }).then(()=>{//6th then
                                    excel.writeRow('FlatFinishActivities',7,{
                                        Activity: arr6[5].milestone,
                                        TowerA : arr6[5].count,
                                        TowerB : arr7[5].count,
                                        TowerC : arr8[5].count,
                                        TowerD : arr9[5].count,
                                        TowerE : arr10[5].count,
                                        TowerF : arr11[5].count,
                                        CompletedPercent : (((parseInt(arr6[5]['count'],10)+parseInt(arr7[5]['count'],10)+parseInt(arr8[5]['count'],10)+parseInt(arr9[5]['count'],10)+parseInt(arr10[5]['count'],10)+parseInt(arr11[5]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                        }).then(()=>{//7th then     
                                        excel.writeRow('FlatFinishActivities',8,{
                                            Activity: arr6[6].milestone,
                                            TowerA : arr6[6].count,
                                            TowerB : arr7[6].count,
                                            TowerC : arr8[6].count,
                                            TowerD : arr9[6].count,
                                            TowerE : arr10[6].count,
                                            TowerF : arr11[6].count,
                                            CompletedPercent : (((parseInt(arr6[6]['count'],10)+parseInt(arr7[6]['count'],10)+parseInt(arr8[6]['count'],10)+parseInt(arr9[6]['count'],10)+parseInt(arr10[6]['count'],10)+parseInt(arr11[6]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                            }).then(()=>{//8th then
                                            excel.writeRow('FlatFinishActivities',9,{
                                                Activity: arr6[7].milestone,
                                                TowerA : arr6[7].count,
                                                TowerB : arr7[7].count,
                                                TowerC : arr8[7].count,
                                                TowerD : arr9[7].count,
                                                TowerE : arr10[7].count,
                                                TowerF : arr11[7].count,
                                                CompletedPercent : (((parseInt(arr6[7]['count'],10)+parseInt(arr7[7]['count'],10)+parseInt(arr8[7]['count'],10)+parseInt(arr9[7]['count'],10)+parseInt(arr10[7]['count'],10)+parseInt(arr11[7]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                }).then(()=>{//9th then
                                                excel.writeRow('FlatFinishActivities',10,{
                                                    Activity: arr6[8].milestone,
                                                    TowerA : arr6[8].count,
                                                    TowerB : arr7[8].count,
                                                    TowerC : arr8[8].count,
                                                    TowerD : arr9[8].count,
                                                    TowerE : arr10[8].count,
                                                    TowerF : arr11[8].count,
                                                    CompletedPercent : (((parseInt(arr6[8]['count'],10)+parseInt(arr7[8]['count'],10)+parseInt(arr8[8]['count'],10)+parseInt(arr9[8]['count'],10)+parseInt(arr10[8]['count'],10)+parseInt(arr11[8]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                    }).then(()=>{//10th then
                                                    excel.writeRow('FlatFinishActivities',11,{
                                                        Activity: arr6[9].milestone,
                                                        TowerA : arr6[9].count,
                                                        TowerB : arr7[9].count,
                                                        TowerC : arr8[9].count,
                                                        TowerD : arr9[9].count,
                                                        TowerE : arr10[9].count,
                                                        TowerF : arr11[9].count,
                                                        CompletedPercent : (((parseInt(arr6[9]['count'],10)+parseInt(arr7[9]['count'],10)+parseInt(arr8[9]['count'],10)+parseInt(arr9[9]['count'],10)+parseInt(arr10[9]['count'],10)+parseInt(arr11[9]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                        }).then(()=>{//11th then
                                                        excel.writeRow('FlatFinishActivities',12,{
                                                            Activity: arr6[10].milestone,
                                                            TowerA : arr6[10].count,
                                                            TowerB : arr7[10].count,
                                                            TowerC : arr8[10].count,
                                                            TowerD : arr9[10].count,
                                                            TowerE : arr10[10].count,
                                                            TowerF : arr11[10].count,
                                                            CompletedPercent : (((parseInt(arr6[10]['count'],10)+parseInt(arr7[10]['count'],10)+parseInt(arr8[10]['count'],10)+parseInt(arr9[10]['count'],10)+parseInt(arr10[10]['count'],10)+parseInt(arr11[10]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                }).then(()=>{//12th then
                                                            excel.writeRow('FlatFinishActivities',13,{
                                                                Activity: arr6[11].milestone,
                                                                TowerA : arr6[11].count,
                                                                TowerB : arr7[11].count,
                                                                TowerC : arr8[11].count,
                                                                TowerD : arr9[11].count,
                                                                TowerE : arr10[11].count,
                                                                TowerF : arr11[11].count,
                                                                CompletedPercent : (((parseInt(arr6[11]['count'],10)+parseInt(arr7[11]['count'],10)+parseInt(arr8[11]['count'],10)+parseInt(arr9[11]['count'],10)+parseInt(arr10[11]['count'],10)+parseInt(arr11[11]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                }).then(()=>{//13th then
                                                                excel.writeRow('FlatFinishActivities',14,{
                                                                    Activity: arr6[12].milestone,
                                                                    TowerA : arr6[12].count,
                                                                    TowerB : arr7[12].count,
                                                                    TowerC : arr8[12].count,
                                                                    TowerD : arr9[12].count,
                                                                    TowerE : arr10[12].count,
                                                                    TowerF : arr11[12].count,
                                                                    CompletedPercent : (((parseInt(arr6[12]['count'],10)+parseInt(arr7[12]['count'],10)+parseInt(arr8[12]['count'],10)+parseInt(arr9[12]['count'],10)+parseInt(arr10[12]['count'],10)+parseInt(arr11[12]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                    }).then(()=>{//14th then
                                                                    excel.writeRow('FlatFinishActivities',15,{
                                                                        Activity: arr6[13].milestone,
                                                                        TowerA : arr6[13].count,
                                                                        TowerB : arr7[13].count,
                                                                        TowerC : arr8[13].count,
                                                                        TowerD : arr9[13].count,
                                                                        TowerE : arr10[13].count,
                                                                        TowerF : arr11[13].count,
                                                                        CompletedPercent : (((parseInt(arr6[13]['count'],10)+parseInt(arr7[13]['count'],10)+parseInt(arr8[13]['count'],10)+parseInt(arr9[13]['count'],10)+parseInt(arr10[13]['count'],10)+parseInt(arr11[13]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                        }).then(()=>{//15th then
                                                                        excel.writeRow('FlatFinishActivities',16,{
                                                                            Activity: arr6[14].milestone,
                                                                            TowerA : arr6[14].count,
                                                                            TowerB : arr7[14].count,
                                                                            TowerC : arr8[14].count,
                                                                            TowerD : arr9[14].count,
                                                                            TowerE : arr10[14].count,
                                                                            TowerF : arr11[14].count,
                                                                            CompletedPercent : (((parseInt(arr6[14]['count'],10)+parseInt(arr7[14]['count'],10)+parseInt(arr8[14]['count'],10)+parseInt(arr9[14]['count'],10)+parseInt(arr10[14]['count'],10)+parseInt(arr11[14]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                            }).then(()=>{//16th
                                                                            excel.writeRow('FlatFinishActivities',17,{
                                                                                Activity: arr6[15].milestone,
                                                                                TowerA : arr6[15].count,
                                                                                TowerB : arr7[15].count,
                                                                                TowerC : arr8[15].count,
                                                                                TowerD : arr9[15].count,
                                                                                TowerE : arr10[15].count,
                                                                                TowerF : arr11[15].count,
                                                                                CompletedPercent : (((parseInt(arr6[15]['count'],10)+parseInt(arr7[15]['count'],10)+parseInt(arr8[15]['count'],10)+parseInt(arr9[15]['count'],10)+parseInt(arr10[15]['count'],10)+parseInt(arr11[15]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                }).then(()=>{//17th
                                                                                excel.writeRow('FlatFinishActivities',18,{
                                                                                    Activity: arr6[16].milestone,
                                                                                    TowerA : arr6[16].count,
                                                                                    TowerB : arr7[16].count,
                                                                                    TowerC : arr8[16].count,
                                                                                    TowerD : arr9[16].count,
                                                                                    TowerE : arr10[16].count,
                                                                                    TowerF : arr11[16].count,
                                                                                    CompletedPercent : (((parseInt(arr6[16]['count'],10)+parseInt(arr7[16]['count'],10)+parseInt(arr8[16]['count'],10)+parseInt(arr9[16]['count'],10)+parseInt(arr10[16]['count'],10)+parseInt(arr11[16]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                    }).then(()=>{//18th
                                                                                    excel.writeRow('FlatFinishActivities',19,{
                                                                                        Activity: arr6[17].milestone,
                                                                                        TowerA : arr6[17].count,
                                                                                        TowerB : arr7[17].count,
                                                                                        TowerC : arr8[17].count,
                                                                                        TowerD : arr9[17].count,
                                                                                        TowerE : arr10[17].count,
                                                                                        TowerF : arr11[17].count,
                                                                                        CompletedPercent : (((parseInt(arr6[17]['count'],10)+parseInt(arr7[17]['count'],10)+parseInt(arr8[17]['count'],10)+parseInt(arr9[17]['count'],10)+parseInt(arr10[17]['count'],10)+parseInt(arr11[17]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                        }).then(()=>{//19th
                                                                                        excel.writeRow('FlatFinishActivities',20,{
                                                                                            Activity: arr6[18].milestone,
                                                                                            TowerA : arr6[18].count,
                                                                                            TowerB : arr7[18].count,
                                                                                            TowerC : arr8[18].count,
                                                                                            TowerD : arr9[18].count,
                                                                                            TowerE : arr10[18].count,
                                                                                            TowerF : arr11[18].count,
                                                                                            CompletedPercent : (((parseInt(arr6[18]['count'],10)+parseInt(arr7[18]['count'],10)+parseInt(arr8[18]['count'],10)+parseInt(arr9[18]['count'],10)+parseInt(arr10[18]['count'],10)+parseInt(arr11[18]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                            }).then(()=>{//20th
                                                                                            excel.writeRow('FlatFinishActivities',21,{
                                                                                                Activity: arr6[19].milestone,
                                                                                                TowerA : arr6[19].count,
                                                                                                TowerB : arr7[19].count,
                                                                                                TowerC : arr8[19].count,
                                                                                                TowerD : arr9[19].count,
                                                                                                TowerE : arr10[19].count,
                                                                                                TowerF : arr11[19].count,
                                                                                                CompletedPercent : (((parseInt(arr6[19]['count'],10)+parseInt(arr7[19]['count'],10)+parseInt(arr8[19]['count'],10)+parseInt(arr9[19]['count'],10)+parseInt(arr10[19]['count'],10)+parseInt(arr11[19]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                }).then(() => {//21st
                                                                                                excel.writeRow('FlatFinishActivities',22,{
                                                                                                    Activity: arr6[20].milestone,
                                                                                                    TowerA : arr6[20].count,
                                                                                                    TowerB : arr7[20].count,
                                                                                                    TowerC : arr8[20].count,
                                                                                                    TowerD : arr9[20].count,
                                                                                                    TowerE : arr10[20].count,
                                                                                                    TowerF : arr11[20].count,
                                                                                                    CompletedPercent : (((parseInt(arr6[20]['count'],10)+parseInt(arr7[20]['count'],10)+parseInt(arr8[20]['count'],10)+parseInt(arr9[20]['count'],10)+parseInt(arr10[20]['count'],10)+parseInt(arr11[20]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                    }).then(() => {//22nd
                                                                                                    excel.writeRow('FlatFinishActivities',23,{
                                                                                                        Activity: arr6[21].milestone,
                                                                                                        TowerA : arr6[21].count,
                                                                                                        TowerB : arr7[21].count,
                                                                                                        TowerC : arr8[21].count,
                                                                                                        TowerD : arr9[21].count,
                                                                                                        TowerE : arr10[21].count,
                                                                                                        TowerF : arr11[21].count,
                                                                                                        CompletedPercent : (((parseInt(arr6[21]['count'],10)+parseInt(arr7[21]['count'],10)+parseInt(arr8[21]['count'],10)+parseInt(arr9[21]['count'],10)+parseInt(arr10[21]['count'],10)+parseInt(arr11[21]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                        }).then(() => {//23rd
                                                                                                        excel.writeRow('FlatFinishActivities',24,{
                                                                                                            Activity: arr6[22].milestone,
                                                                                                            TowerA : arr6[22].count,
                                                                                                            TowerB : arr7[22].count,
                                                                                                            TowerC : arr8[22].count,
                                                                                                            TowerD : arr9[22].count,
                                                                                                            TowerE : arr10[22].count,
                                                                                                            TowerF : arr11[22].count,
                                                                                                            CompletedPercent : (((parseInt(arr6[22]['count'],10)+parseInt(arr7[22]['count'],10)+parseInt(arr8[22]['count'],10)+parseInt(arr9[22]['count'],10)+parseInt(arr10[22]['count'],10)+parseInt(arr11[22]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                            }).then(() => {//24th
                                                                                                            excel.writeRow('FlatFinishActivities',25,{
                                                                                                                Activity: arr6[23].milestone,
                                                                                                                TowerA : arr6[23].count,
                                                                                                                TowerB : arr7[23].count,
                                                                                                                TowerC : arr8[23].count,
                                                                                                                TowerD : arr9[23].count,
                                                                                                                TowerE : arr10[23].count,
                                                                                                                TowerF : arr11[23].count,
                                                                                                                CompletedPercent : (((parseInt(arr6[23]['count'],10)+parseInt(arr7[23]['count'],10)+parseInt(arr8[23]['count'],10)+parseInt(arr9[23]['count'],10)+parseInt(arr10[23]['count'],10)+parseInt(arr11[23]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                }).then(() => {//25th
                                                                                                                excel.writeRow('FlatFinishActivities',26,{
                                                                                                                    Activity: arr6[24].milestone,
                                                                                                                    TowerA : arr6[24].count,
                                                                                                                    TowerB : arr7[24].count,
                                                                                                                    TowerC : arr8[24].count,
                                                                                                                    TowerD : arr9[24].count,
                                                                                                                    TowerE : arr10[24].count,
                                                                                                                    TowerF : arr11[24].count,
                                                                                                                    CompletedPercent : (((parseInt(arr6[24]['count'],10)+parseInt(arr7[24]['count'],10)+parseInt(arr8[24]['count'],10)+parseInt(arr9[24]['count'],10)+parseInt(arr10[24]['count'],10)+parseInt(arr11[24]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                     }).then(() =>{//26th
                                                                                                                    excel.writeRow('FlatFinishActivities',27,{
                                                                                                                        Activity: arr6[25].milestone,
                                                                                                                        TowerA : arr6[25].count,
                                                                                                                        TowerB : arr7[25].count,
                                                                                                                        TowerC : arr8[25].count,
                                                                                                                        TowerD : arr9[25].count,
                                                                                                                        TowerE : arr10[25].count,
                                                                                                                        TowerF : arr11[25].count,
                                                                                                                        CompletedPercent : (((parseInt(arr6[25]['count'],10)+parseInt(arr7[25]['count'],10)+parseInt(arr8[25]['count'],10)+parseInt(arr9[25]['count'],10)+parseInt(arr10[25]['count'],10)+parseInt(arr11[25]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                         }).then(()=>{//27th
                                                                                                                        excel.writeRow('FlatFinishActivities',28,{
                                                                                                                            Activity: arr6[26].milestone,
                                                                                                                            TowerA : arr6[26].count,
                                                                                                                            TowerB : arr7[26].count,
                                                                                                                            TowerC : arr8[26].count,
                                                                                                                            TowerD : arr9[26].count,
                                                                                                                            TowerE : arr10[26].count,
                                                                                                                            TowerF : arr11[26].count,
                                                                                                                            CompletedPercent : (((parseInt(arr6[26]['count'],10)+parseInt(arr7[26]['count'],10)+parseInt(arr8[26]['count'],10)+parseInt(arr9[26]['count'],10)+parseInt(arr10[26]['count'],10)+parseInt(arr11[26]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                        }).then(() =>{//28th
                                                                                                                        excel.writeRow('FlatFinishActivities',29,{
                                                                                                                            Activity: arr6[27].milestone,
                                                                                                                            TowerA : arr6[27].count,
                                                                                                                            TowerB : arr7[27].count,
                                                                                                                            TowerC : arr8[27].count,
                                                                                                                            TowerD : arr9[27].count,
                                                                                                                            TowerE : arr10[27].count,
                                                                                                                            TowerF : arr11[27].count,
                                                                                                                            CompletedPercent : (((parseInt(arr6[27]['count'],10)+parseInt(arr7[27]['count'],10)+parseInt(arr8[27]['count'],10)+parseInt(arr9[27]['count'],10)+parseInt(arr10[27]['count'],10)+parseInt(arr11[27]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                  }).then(()=>{//29th
                                                                                                                    excel.writeRow('FlatFinishActivities',30,{
                                                                                                                        Activity: arr6[28].milestone,
                                                                                                                        TowerA : arr6[28].count,
                                                                                                                        TowerB : arr7[28].count,
                                                                                                                        TowerC : arr8[28].count,
                                                                                                                        TowerD : arr9[28].count,
                                                                                                                        TowerE : arr10[28].count,
                                                                                                                        TowerF : arr11[28].count,
                                                                                                                        CompletedPercent : (((parseInt(arr6[28]['count'],10)+parseInt(arr7[28]['count'],10)+parseInt(arr8[28]['count'],10)+parseInt(arr9[28]['count'],10)+parseInt(arr10[28]['count'],10)+parseInt(arr11[28]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                }).then(()=>{//30th
                                                                                                                        excel.writeRow('FlatFinishActivities',31,{
                                                                                                                            Activity: arr6[29].milestone,
                                                                                                                            TowerA : arr6[29].count,
                                                                                                                            TowerB : arr7[29].count,
                                                                                                                            TowerC : arr8[29].count,
                                                                                                                            TowerD : arr9[29].count,
                                                                                                                            TowerE : arr10[29].count,
                                                                                                                            TowerF : arr11[29].count,
                                                                                                                            CompletedPercent : (((parseInt(arr6[29]['count'],10)+parseInt(arr7[29]['count'],10)+parseInt(arr8[29]['count'],10)+parseInt(arr9[29]['count'],10)+parseInt(arr10[29]['count'],10)+parseInt(arr11[29]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                            }).then(()=>{//31st
                                                                                                                            excel.writeRow('FlatFinishActivities',32,{
                                                                                                                                Activity: arr6[30].milestone,
                                                                                                                                TowerA : arr6[30].count,
                                                                                                                                TowerB : arr7[30].count,
                                                                                                                                TowerC : arr8[30].count,
                                                                                                                                TowerD : arr9[30].count,
                                                                                                                                TowerE : arr10[30].count,
                                                                                                                                TowerF : arr11[30].count,
                                                                                                                                CompletedPercent : (((parseInt(arr6[30]['count'],10)+parseInt(arr7[30]['count'],10)+parseInt(arr8[30]['count'],10)+parseInt(arr9[30]['count'],10)+parseInt(arr10[30]['count'],10)+parseInt(arr11[30]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                }).then(() =>{//32nd
                                                                                                                                excel.writeRow('FlatFinishActivities',33,{
                                                                                                                                    Activity: arr6[31].milestone,
                                                                                                                                    TowerA : arr6[31].count,
                                                                                                                                    TowerB : arr7[31].count,
                                                                                                                                    TowerC : arr8[31].count,
                                                                                                                                    TowerD : arr9[31].count,
                                                                                                                                    TowerE : arr10[31].count,
                                                                                                                                    TowerF : arr11[31].count,
                                                                                                                                    CompletedPercent : (((parseInt(arr6[31]['count'],10)+parseInt(arr7[31]['count'],10)+parseInt(arr8[31]['count'],10)+parseInt(arr9[31]['count'],10)+parseInt(arr10[31]['count'],10)+parseInt(arr11[31]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                    }).then(()=>{//33rd
                                                                                                                                    excel.writeRow('FlatFinishActivities',34,{
                                                                                                                                        Activity: arr6[32].milestone,
                                                                                                                                        TowerA : arr6[32].count,
                                                                                                                                        TowerB : arr7[32].count,
                                                                                                                                        TowerC : arr8[32].count,
                                                                                                                                        TowerD : arr9[32].count,
                                                                                                                                        TowerE : arr10[32].count,
                                                                                                                                        TowerF : arr11[32].count,
                                                                                                                                        CompletedPercent : (((parseInt(arr6[32]['count'],10)+parseInt(arr7[32]['count'],10)+parseInt(arr8[32]['count'],10)+parseInt(arr9[32]['count'],10)+parseInt(arr10[32]['count'],10)+parseInt(arr11[32]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                        }).then(() =>{//34th
                                                                                                                                        excel.writeRow('FlatFinishActivities',35,{
                                                                                                                                            Activity: arr6[33].milestone,
                                                                                                                                            TowerA : arr6[33].count,
                                                                                                                                            TowerB : arr7[33].count,
                                                                                                                                            TowerC : arr8[33].count,
                                                                                                                                            TowerD : arr9[33].count,
                                                                                                                                            TowerE : arr10[33].count,
                                                                                                                                            TowerF : arr11[33].count,
                                                                                                                                            CompletedPercent : (((parseInt(arr6[33]['count'],10)+parseInt(arr7[33]['count'],10)+parseInt(arr8[33]['count'],10)+parseInt(arr9[33]['count'],10)+parseInt(arr10[33]['count'],10)+parseInt(arr11[33]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                            }).then(() => {//35th
                                                                                                                                            excel.writeRow('FlatFinishActivities',36,{
                                                                                                                                                Activity: arr6[34].milestone,
                                                                                                                                                TowerA : arr6[34].count,
                                                                                                                                                TowerB : arr7[34].count,
                                                                                                                                                TowerC : arr8[34].count,
                                                                                                                                                TowerD : arr9[34].count,
                                                                                                                                                TowerE : arr10[34].count,
                                                                                                                                                TowerF : arr11[34].count,
                                                                                                                                                CompletedPercent : (((parseInt(arr6[34]['count'],10)+parseInt(arr7[34]['count'],10)+parseInt(arr8[34]['count'],10)+parseInt(arr9[34]['count'],10)+parseInt(arr10[34]['count'],10)+parseInt(arr11[34]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                               }).then(() => {//36th
                                                                                                                                                excel.writeRow('FlatFinishActivities',37,{
                                                                                                                                                    Activity: arr6[35].milestone,
                                                                                                                                                    TowerA : arr6[35].count,
                                                                                                                                                    TowerB : arr7[35].count,
                                                                                                                                                    TowerC : arr8[35].count,
                                                                                                                                                    TowerD : arr9[35].count,
                                                                                                                                                    TowerE : arr10[35].count,
                                                                                                                                                    TowerF : arr11[35].count,
                                                                                                                                                    CompletedPercent : (((parseInt(arr6[35]['count'],10)+parseInt(arr7[35]['count'],10)+parseInt(arr8[35]['count'],10)+parseInt(arr9[35]['count'],10)+parseInt(arr10[35]['count'],10)+parseInt(arr11[35]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                                    })
                                                                                                                                            })//36th
                                                                                                                                        })//35th    
                                                                                                                                    })//34th
                                                                                                                                })//33rd
                                                                                                                            })//32nd
                                                                                                                        })//31st
                                                                                                                    })//30th
                                                                                                                  })//29th
                                                                                                                })//28th
                                                                                                              })//27th
                                                                                                            })//26th
                                                                                                          })//25th
                                                                                                        })//24th
                                                                                                    })//23rd 
                                                                                                })//22nd
                                                                                            })//21st
                                                                                        })//20th
                                                                                    })//19th then
                                                                                })//18th then 
                                                                            })//17th then
                                                                        })//16th then
                                                                    })//15th then ending
                                                                })//14th then ending
                                                            })//13th then ending
                                                        })//12th then ending
                                                    })//11th then ending
                                                })//10th then ending
                                            })//9th then ending
                                        })//8th then ending
                                    })//7th then ending
                                });//6th then ending
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;        
}


//flat finish excel fror phase 2
function  flatfinishExcelGreenfield2(floor_count,arr6,arr7,arr8)
{

            //path to generate excel
            const path = require('path');
            let excel = new Excel(path.join(__dirname,'GreenfieldPhase2.xlsx'));
    
            excel.writeSheet('FlatFinishActivities',['Activity','TowerG','TowerH','TowerJ','CompletedPercent'],[floor_count]).then(()=>{
                excel.writeRow('FlatFinishActivities',1,{
                   Activity: "Floor Count",
                   TowerG : floor_count[0]['total_unit'],
                   TowerH : floor_count[1]['total_unit'],
                   TowerJ : floor_count[2]['total_unit'],
               }).then(()=>{//1st then
                   excel.writeRow('FlatFinishActivities',2,{
                    Activity: arr6[0].milestone,
                    TowerG : arr6[0].count,
                    TowerH : arr7[0].count,
                    TowerJ : arr8[0].count,
                    CompletedPercent : (((parseInt(arr6[0]['count'],10)+parseInt(arr7[0]['count'],10)+parseInt(arr8[0]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                 }).then(()=>{//2nd then
                    excel.writeRow('FlatFinishActivities',3,{
                        Activity: arr6[1].milestone,
                        TowerG : arr6[1].count,
                        TowerH : arr7[1].count,
                        TowerJ : arr8[1].count,
                        CompletedPercent : (((parseInt(arr6[1]['count'],10)+parseInt(arr7[1]['count'],10)+parseInt(arr8[1]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                        }).then(()=>{//3rd then
                        excel.writeRow('FlatFinishActivities',4,{
                            Activity: arr6[2].milestone,
                            TowerG : arr6[2].count,
                            TowerH : arr7[2].count,
                            TowerJ : arr8[2].count,
                            CompletedPercent : (((parseInt(arr6[2]['count'],10)+parseInt(arr7[2]['count'],10)+parseInt(arr8[2]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                            }).then(()=>{//4th then
                            excel.writeRow('FlatFinishActivities',5,{
                                Activity: arr6[3].milestone,
                                TowerG : arr6[3].count,
                                TowerH : arr7[3].count,
                                TowerJ : arr8[3].count,
                                CompletedPercent : (((parseInt(arr6[3]['count'],10)+parseInt(arr7[3]['count'],10)+parseInt(arr8[3]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                }).then(()=>{//5th then
                                excel.writeRow('FlatFinishActivities',6,{
                                    Activity: arr6[4].milestone,
                                    TowerG : arr6[4].count,
                                    TowerH : arr7[4].count,
                                    TowerJ : arr8[4].count,
                                    CompletedPercent : (((parseInt(arr6[4]['count'],10)+parseInt(arr7[4]['count'],10)+parseInt(arr8[4]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                    }).then(()=>{//6th then
                                    excel.writeRow('FlatFinishActivities',7,{
                                        Activity: arr6[5].milestone,
                                        TowerG : arr6[5].count,
                                        TowerH : arr7[5].count,
                                        TowerJ : arr8[5].count,
                                        CompletedPercent : (((parseInt(arr6[5]['count'],10)+parseInt(arr7[5]['count'],10)+parseInt(arr8[5]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                        }).then(()=>{//7th then     
                                        excel.writeRow('FlatFinishActivities',8,{
                                            Activity: arr6[6].milestone,
                                            TowerG : arr6[6].count,
                                            TowerH : arr7[6].count,
                                            TowerJ : arr8[6].count,
                                            CompletedPercent : (((parseInt(arr6[6]['count'],10)+parseInt(arr7[6]['count'],10)+parseInt(arr8[6]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                            }).then(()=>{//8th then
                                            excel.writeRow('FlatFinishActivities',9,{
                                                Activity: arr6[7].milestone,
                                                TowerG : arr6[7].count,
                                                TowerH : arr7[7].count,
                                                TowerJ : arr8[7].count,
                                                CompletedPercent : (((parseInt(arr6[7]['count'],10)+parseInt(arr7[7]['count'],10)+parseInt(arr8[7]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                }).then(()=>{//9th then
                                                excel.writeRow('FlatFinishActivities',10,{
                                                    Activity: arr6[8].milestone,
                                                    TowerG : arr6[8].count,
                                                    TowerH : arr7[8].count,
                                                    TowerJ : arr8[8].count,
                                                    CompletedPercent : (((parseInt(arr6[8]['count'],10)+parseInt(arr7[8]['count'],10)+parseInt(arr8[8]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                    }).then(()=>{//10th then
                                                    excel.writeRow('FlatFinishActivities',11,{
                                                        Activity: arr6[9].milestone,
                                                        TowerG : arr6[9].count,
                                                        TowerH : arr7[9].count,
                                                        TowerJ : arr8[9].count,
                                                        CompletedPercent : (((parseInt(arr6[9]['count'],10)+parseInt(arr7[9]['count'],10)+parseInt(arr8[9]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                        }).then(()=>{//11th then
                                                        excel.writeRow('FlatFinishActivities',12,{
                                                            Activity: arr6[10].milestone,
                                                            TowerG : arr6[10].count,
                                                            TowerH : arr7[10].count,
                                                            TowerJ : arr8[10].count,
                                                            CompletedPercent : (((parseInt(arr6[10]['count'],10)+parseInt(arr7[10]['count'],10)+parseInt(arr8[10]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                            }).then(()=>{//12th then
                                                            excel.writeRow('FlatFinishActivities',13,{
                                                                Activity: arr6[11].milestone,
                                                                TowerG : arr6[11].count,
                                                                TowerH : arr7[11].count,
                                                                TowerJ : arr8[11].count,
                                                                CompletedPercent : (((parseInt(arr6[11]['count'],10)+parseInt(arr7[11]['count'],10)+parseInt(arr8[11]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                }).then(()=>{//13th then
                                                                excel.writeRow('FlatFinishActivities',14,{
                                                                    Activity: arr6[12].milestone,
                                                                    TowerG : arr6[12].count,
                                                                    TowerH : arr7[12].count,
                                                                    TowerJ : arr8[12].count,
                                                                    CompletedPercent : (((parseInt(arr6[12]['count'],10)+parseInt(arr7[12]['count'],10)+parseInt(arr8[12]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                    }).then(()=>{//14th then
                                                                    excel.writeRow('FlatFinishActivities',15,{
                                                                        Activity: arr6[13].milestone,
                                                                        TowerG : arr6[13].count,
                                                                        TowerH : arr7[13].count,
                                                                        TowerJ : arr8[13].count,
                                                                        CompletedPercent : (((parseInt(arr6[13]['count'],10)+parseInt(arr7[13]['count'],10)+parseInt(arr8[13]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                        }).then(()=>{//15th then
                                                                        excel.writeRow('FlatFinishActivities',16,{
                                                                            Activity: arr6[14].milestone,
                                                                            TowerG : arr6[14].count,
                                                                            TowerH : arr7[14].count,
                                                                            TowerJ : arr8[14].count,
                                                                            CompletedPercent : (((parseInt(arr6[14]['count'],10)+parseInt(arr7[14]['count'],10)+parseInt(arr8[14]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                            }).then(()=>{//16th
                                                                            excel.writeRow('FlatFinishActivities',17,{
                                                                                Activity: arr6[15].milestone,
                                                                                TowerG : arr6[15].count,
                                                                                TowerH : arr7[15].count,
                                                                                TowerJ : arr8[15].count,
                                                                                CompletedPercent : (((parseInt(arr6[15]['count'],10)+parseInt(arr7[15]['count'],10)+parseInt(arr8[15]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                }).then(()=>{//17th
                                                                                excel.writeRow('FlatFinishActivities',18,{
                                                                                    Activity: arr6[16].milestone,
                                                                                    TowerG : arr6[16].count,
                                                                                    TowerH : arr7[16].count,
                                                                                    TowerJ : arr8[16].count,
                                                                                    CompletedPercent : (((parseInt(arr6[16]['count'],10)+parseInt(arr7[16]['count'],10)+parseInt(arr8[16]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                    }).then(()=>{//18th
                                                                                    excel.writeRow('FlatFinishActivities',19,{
                                                                                        Activity: arr6[17].milestone,
                                                                                        TowerG : arr6[17].count,
                                                                                        TowerH : arr7[17].count,
                                                                                        TowerJ : arr8[17].count,
                                                                                        CompletedPercent : (((parseInt(arr6[17]['count'],10)+parseInt(arr7[17]['count'],10)+parseInt(arr8[17]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                        }).then(()=>{//19th
                                                                                        excel.writeRow('FlatFinishActivities',20,{
                                                                                            Activity: arr6[18].milestone,
                                                                                            TowerG : arr6[18].count,
                                                                                            TowerH : arr7[18].count,
                                                                                            TowerJ : arr8[18].count,
                                                                                            CompletedPercent : (((parseInt(arr6[18]['count'],10)+parseInt(arr7[18]['count'],10)+parseInt(arr8[18]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                            }).then(()=>{//20th
                                                                                            excel.writeRow('FlatFinishActivities',21,{
                                                                                                Activity: arr6[19].milestone,
                                                                                                TowerG : arr6[19].count,
                                                                                                TowerH : arr7[19].count,
                                                                                                TowerJ : arr8[19].count,
                                                                                                CompletedPercent : (((parseInt(arr6[19]['count'],10)+parseInt(arr7[19]['count'],10)+parseInt(arr8[19]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                }).then(() => {//21st
                                                                                                excel.writeRow('FlatFinishActivities',22,{
                                                                                                    Activity: arr6[20].milestone,
                                                                                                    TowerG : arr6[20].count,
                                                                                                    TowerH : arr7[20].count,
                                                                                                    TowerJ : arr8[20].count,
                                                                                                    CompletedPercent : (((parseInt(arr6[20]['count'],10)+parseInt(arr7[20]['count'],10)+parseInt(arr8[20]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                    }).then(() => {//22nd
                                                                                                    excel.writeRow('FlatFinishActivities',23,{
                                                                                                        Activity: arr6[21].milestone,
                                                                                                        TowerG : arr6[21].count,
                                                                                                        TowerH : arr7[21].count,
                                                                                                        TowerJ : arr8[21].count,
                                                                                                        CompletedPercent : (((parseInt(arr6[21]['count'],10)+parseInt(arr7[21]['count'],10)+parseInt(arr8[21]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                        }).then(() => {//23rd
                                                                                                        excel.writeRow('FlatFinishActivities',24,{
                                                                                                            Activity: arr6[22].milestone,
                                                                                                            TowerG : arr6[22].count,
                                                                                                            TowerH : arr7[22].count,
                                                                                                            TowerJ : arr8[22].count,
                                                                                                            CompletedPercent : (((parseInt(arr6[22]['count'],10)+parseInt(arr7[22]['count'],10)+parseInt(arr8[22]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                            }).then(() => {//24th
                                                                                                            excel.writeRow('FlatFinishActivities',25,{
                                                                                                                Activity: arr6[23].milestone,
                                                                                                                TowerG : arr6[23].count,
                                                                                                                TowerH : arr7[23].count,
                                                                                                                TowerJ : arr8[23].count,
                                                                                                                CompletedPercent : (((parseInt(arr6[23]['count'],10)+parseInt(arr7[23]['count'],10)+parseInt(arr8[23]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                }).then(() => {//25th
                                                                                                                excel.writeRow('FlatFinishActivities',26,{
                                                                                                                    Activity: arr6[24].milestone,
                                                                                                                    TowerG : arr6[24].count,
                                                                                                                    TowerH : arr7[24].count,
                                                                                                                    TowerJ : arr8[24].count,
                                                                                                                    CompletedPercent : (((parseInt(arr6[24]['count'],10)+parseInt(arr7[24]['count'],10)+parseInt(arr8[24]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                     }).then(() =>{//26th
                                                                                                                    excel.writeRow('FlatFinishActivities',27,{
                                                                                                                        Activity: arr6[25].milestone,
                                                                                                                        TowerG : arr6[25].count,
                                                                                                                        TowerH : arr7[25].count,
                                                                                                                        TowerJ : arr8[25].count,
                                                                                                                        CompletedPercent : (((parseInt(arr6[25]['count'],10)+parseInt(arr7[25]['count'],10)+parseInt(arr8[25]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                         }).then(()=>{//27th
                                                                                                                        excel.writeRow('FlatFinishActivities',28,{
                                                                                                                            Activity: arr6[26].milestone,
                                                                                                                            TowerG : arr6[26].count,
                                                                                                                            TowerH : arr7[26].count,
                                                                                                                            TowerJ : arr8[26].count,
                                                                                                                            CompletedPercent : (((parseInt(arr6[26]['count'],10)+parseInt(arr7[26]['count'],10)+parseInt(arr8[26]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                        }).then(() =>{//28th
                                                                                                                        excel.writeRow('FlatFinishActivities',29,{
                                                                                                                            Activity: arr6[27].milestone,
                                                                                                                            TowerG : arr6[27].count,
                                                                                                                            TowerH : arr7[27].count,
                                                                                                                            TowerJ : arr8[27].count,
                                                                                                                            CompletedPercent : (((parseInt(arr6[27]['count'],10)+parseInt(arr7[27]['count'],10)+parseInt(arr8[27]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                  }).then(()=>{//29th
                                                                                                                    excel.writeRow('FlatFinishActivities',30,{
                                                                                                                        Activity: arr6[28].milestone,
                                                                                                                        TowerG : arr6[28].count,
                                                                                                                        TowerH : arr7[28].count,
                                                                                                                        TowerJ : arr8[28].count,
                                                                                                                        CompletedPercent : (((parseInt(arr6[28]['count'],10)+parseInt(arr7[28]['count'],10)+parseInt(arr8[28]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                }).then(()=>{//30th
                                                                                                                        excel.writeRow('FlatFinishActivities',31,{
                                                                                                                            Activity: arr6[29].milestone,
                                                                                                                            TowerG : arr6[29].count,
                                                                                                                            TowerH : arr7[29].count,
                                                                                                                            TowerJ : arr8[29].count,
                                                                                                                            CompletedPercent : (((parseInt(arr6[29]['count'],10)+parseInt(arr7[29]['count'],10)+parseInt(arr8[29]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                            }).then(()=>{//31st
                                                                                                                            excel.writeRow('FlatFinishActivities',32,{
                                                                                                                                Activity: arr6[30].milestone,
                                                                                                                                TowerG : arr6[30].count,
                                                                                                                                TowerH : arr7[30].count,
                                                                                                                                TowerJ : arr8[30].count,
                                                                                                                                CompletedPercent : (((parseInt(arr6[30]['count'],10)+parseInt(arr7[30]['count'],10)+parseInt(arr8[30]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                }).then(() =>{//32nd
                                                                                                                                excel.writeRow('FlatFinishActivities',33,{
                                                                                                                                    Activity: arr6[31].milestone,
                                                                                                                                    TowerG : arr6[31].count,
                                                                                                                                    TowerH : arr7[31].count,
                                                                                                                                    TowerJ : arr8[31].count,
                                                                                                                                    CompletedPercent : (((parseInt(arr6[31]['count'],10)+parseInt(arr7[31]['count'],10)+parseInt(arr8[31]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                    }).then(()=>{//33rd
                                                                                                                                    excel.writeRow('FlatFinishActivities',34,{
                                                                                                                                        Activity: arr6[32].milestone,
                                                                                                                                        TowerG : arr6[32].count,
                                                                                                                                        TowerH : arr7[32].count,
                                                                                                                                        TowerJ : arr8[32].count,
                                                                                                                                        CompletedPercent : (((parseInt(arr6[32]['count'],10)+parseInt(arr7[32]['count'],10)+parseInt(arr8[32]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                        }).then(() =>{//34th
                                                                                                                                        excel.writeRow('FlatFinishActivities',35,{
                                                                                                                                            Activity: arr6[33].milestone,
                                                                                                                                            TowerG : arr6[33].count,
                                                                                                                                            TowerH : arr7[33].count,
                                                                                                                                            TowerJ : arr8[33].count,
                                                                                                                                            CompletedPercent : (((parseInt(arr6[33]['count'],10)+parseInt(arr7[33]['count'],10)+parseInt(arr8[33]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                            }).then(() => {//35th
                                                                                                                                            excel.writeRow('FlatFinishActivities',36,{
                                                                                                                                                Activity: arr6[34].milestone,
                                                                                                                                                TowerG : arr6[34].count,
                                                                                                                                                TowerH : arr7[34].count,
                                                                                                                                                TowerJ : arr8[34].count,
                                                                                                                                                CompletedPercent : (((parseInt(arr6[34]['count'],10)+parseInt(arr7[34]['count'],10)+parseInt(arr8[34]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                               }).then(() => {//36th
                                                                                                                                                excel.writeRow('FlatFinishActivities',37,{
                                                                                                                                                    Activity: arr6[35].milestone,
                                                                                                                                                    TowerG : arr6[35].count,
                                                                                                                                                    TowerH : arr7[35].count,
                                                                                                                                                    TowerJ : arr8[35].count,
                                                                                                                                                    CompletedPercent : (((parseInt(arr6[35]['count'],10)+parseInt(arr7[35]['count'],10)+parseInt(arr8[35]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                                    })
                                                                                                                                            })//36th
                                                                                                                                        })//35th    
                                                                                                                                    })//34th
                                                                                                                                })//33rd
                                                                                                                            })//32nd
                                                                                                                        })//31st
                                                                                                                    })//30th
                                                                                                                  })//29th
                                                                                                                })//28th
                                                                                                              })//27th
                                                                                                            })//26th
                                                                                                          })//25th
                                                                                                        })//24th
                                                                                                    })//23rd 
                                                                                                })//22nd
                                                                                            })//21st
                                                                                        })//20th
                                                                                    })//19th then
                                                                                })//18th then 
                                                                            })//17th then
                                                                        })//16th then
                                                                    })//15th then ending
                                                                })//14th then ending
                                                            })//13th then ending
                                                        })//12th then ending
                                                    })//11th then ending
                                                })//10th then ending
                                            })//9th then ending
                                        })//8th then ending
                                    })//7th then ending
                                });//6th then ending
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;        
}

//function to generate excel
function generateExcel(arr,arr2,arr3,arr4,arr5)
{

    //path to generate excel
    const path = require('path');
    let excel = new Excel(path.join(__dirname,'PanoramaHills.xlsx'));

        // var tower = arr[0]['tower_name'];
       excel.writeSheet('StructureActivities',['Activity','Tower2A','Tower2B','Tower2C','Tower2D','CompletedPercent'],[arr]).then(()=>{
                 excel.writeRow('StructureActivities',1,{
                    Activity: "Unit Count",
                    Tower2A : arr[0]['total_unit'],
                    Tower2B : arr[1]['total_unit'],
                    Tower2C : arr[2]['total_unit'],
                    Tower2D : arr[3]['total_unit'],
                }).then(()=>{
                        excel.writeRow('StructureActivities',2,{
                            Activity : arr2[0].milestone,
                            Tower2A  : arr2[0].count,
                            Tower2B  : arr3[0].count,
                            Tower2C  : arr4[0].count,
                            Tower2D  : arr5[0].count,
                            CompletedPercent : ((arr2[0]['count']+arr3[0]['count']+arr4[0]['count']+arr5[0]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"
                        }).then(()=>{
                        excel.writeRow('StructureActivities',3,{
                            Activity : arr2[1].milestone,
                            Tower2A :arr2[1].count,
                            Tower2B :arr3[1].count,
                            Tower2C  : arr4[1].count,
                            Tower2D  : arr5[1].count,
                            CompletedPercent : ((arr2[1]['count']+arr3[1]['count']+arr4[1]['count']+arr5[1]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"
                        }).then(()=>{
                            excel.writeRow('StructureActivities',4,{
                                Activity : arr2[2].milestone,
                                Tower2A :  arr2[2].count,
                                Tower2B :  arr3[2].count,
                                Tower2C  : arr4[2].count,
                                Tower2D  : arr5[2].count,
                                CompletedPercent : ((arr2[2]['count']+arr3[2]['count']+arr4[2]['count']+arr5[2]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                    
                            }).then(()=>{
                                excel.writeRow('StructureActivities',5,{
                                    Activity : arr2[3].milestone,
                                    Tower2A :arr2[3].count,
                                    Tower2B :arr3[3].count,
                                    Tower2C  : arr4[3].count,
                                    Tower2D  : arr5[3].count,
                                    CompletedPercent : ((arr2[3]['count']+arr3[3]['count']+arr4[3]['count']+arr5[3]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                           
                                }).then(()=>{
                                    excel.writeRow('StructureActivities',6,{
                                      Activity : arr2[4].milestone, 
                                      Tower2A :arr2[4].count,
                                      Tower2B :arr3[4].count,
                                      Tower2C  : arr4[4].count,
                                      Tower2D  : arr5[4].count,
                                      CompletedPercent : ((arr2[4]['count']+arr3[4]['count']+arr4[4]['count']+arr5[4]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                    }).then(()=>{
                                        excel.writeRow('StructureActivities',7,{
                                            Activity : arr2[5].milestone,
                                            Tower2A :arr2[5].count,
                                            Tower2B :arr3[5].count,
                                            Tower2C  : arr4[5].count,
                                            Tower2D  : arr5[5].count,
                                            CompletedPercent : ((arr2[5]['count']+arr3[5]['count']+arr4[5]['count']+arr5[5]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                                                                 
                                          }).then(()=>{
                                            excel.writeRow('StructureActivities',8,{
                                                Activity : arr2[6].milestone,
                                                Tower2A :arr2[6].count,
                                                Tower2B :arr3[6].count,
                                                Tower2C  : arr4[6].count,
                                                Tower2D  : arr5[6].count,
                                                CompletedPercent : ((arr2[6]['count']+arr3[6]['count']+arr4[6]['count']+arr5[6]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                                                                           
                                              }).then(()=>{
                                                excel.writeRow('StructureActivities',9,{
                                                    Activity : arr2[7].milestone,
                                                    Tower2A :arr2[7].count,
                                                    Tower2B :arr3[7].count,
                                                    Tower2C  : arr4[7].count,
                                                    Tower2D  : arr5[7].count,
                                                    CompletedPercent : ((arr2[7]['count']+arr3[7]['count']+arr4[7]['count']+arr5[7]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                            
                                                  }).then(()=>{
                                                    excel.writeRow('StructureActivities',10,{
                                                        Activity : arr2[8].milestone,
                                                        Tower2A :arr2[8].count,
                                                        Tower2B :arr3[8].count,
                                                        Tower2C  : arr4[8].count,
                                                        Tower2D  : arr5[8].count,
                                                        CompletedPercent : ((arr2[8]['count']+arr3[8]['count']+arr4[8]['count']+arr5[8]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
              
                                                      }).then(()=>{
                                                        excel.writeRow('StructureActivities',11,{
                                                            Activity : arr2[9].milestone,
                                                            Tower2A :arr2[9].count,
                                                            Tower2B :arr3[9].count,
                                                            Tower2C  : arr4[9].count,
                                                            Tower2D  : arr5[9].count,
                                                            CompletedPercent : ((arr2[9]['count']+arr3[9]['count']+arr4[9]['count']+arr5[9]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                   
                                                          }).then(()=>{
                                                            excel.writeRow('StructureActivities',12,{
                                                                Activity : arr2[10].milestone,
                                                                Tower2A :arr2[10].count,
                                                                Tower2B :arr3[10].count,
                                                                Tower2C  : arr4[10].count,
                                                                Tower2D  : arr5[10].count,
                                                                CompletedPercent : ((arr2[10]['count']+arr3[10]['count']+arr4[10]['count']+arr5[10]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                                                                                                            
                                                            }).then(()=>{
                                                                excel.writeRow('StructureActivities',13,{
                                                                    Activity : arr2[11].milestone,
                                                                    Tower2A :arr2[11].count,
                                                                    Tower2B :arr3[11].count,
                                                                    Tower2C  : arr4[11].count,
                                                                    Tower2D  : arr5[11].count,
                                                                    CompletedPercent : ((arr2[11]['count']+arr3[11]['count']+arr4[11]['count']+arr5[11]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                                                                                                                
                                                                }).then(()=>{
                                                                    excel.writeRow('StructureActivities',14,{
                                                                        Activity : arr2[12].milestone,
                                                                        Tower2A :arr2[12].count,
                                                                        Tower2B :arr3[12].count,
                                                                        Tower2C  : arr4[12].count,
                                                                        Tower2D  : arr5[12].count,
                                                                        CompletedPercent : ((arr2[12]['count']+arr3[12]['count']+arr4[12]['count']+arr5[12]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                                                                                                                    
                                                                    }).then(()=>{
                                                                        excel.writeRow('StructureActivities',15,{
                                                                            Activity : arr2[13].milestone,
                                                                            Tower2A :arr2[13].count,
                                                                            Tower2B :arr3[13].count,
                                                                            Tower2C  : arr4[13].count,
                                                                            Tower2D  : arr5[13].count,
                                                                            CompletedPercent : ((arr2[13]['count']+arr3[13]['count']+arr4[13]['count']+arr5[13]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                   
                                                                        }).then(()=>{
                                                                            excel.writeRow('StructureActivities',16,{
                                                                                Activity : arr2[14].milestone,
                                                                                Tower2A :arr2[14].count,
                                                                                Tower2B :arr3[14].count,
                                                                                Tower2C  : arr4[14].count,
                                                                                Tower2D  : arr5[14].count,
                                                                                CompletedPercent : ((arr2[14]['count']+arr3[14]['count']+arr4[14]['count']+arr5[14]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                       
                                                                            }).then(()=>{
                                                                                excel.writeRow('StructureActivities',17,{
                                                                                    Activity : arr2[15].milestone,
                                                                                    Tower2A :arr2[15].count,
                                                                                    Tower2B :arr3[15].count,
                                                                                    Tower2C  : arr4[15].count,
                                                                                    Tower2D  : arr5[15].count,
                                                                                    CompletedPercent : ((arr2[15]['count']+arr3[15]['count']+arr4[15]['count']+arr5[15]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                             
                                                                                }).then(()=>{
                                                                                    excel.writeRow('StructureActivities',18,{
                                                                                        Activity : arr2[16].milestone,
                                                                                        Tower2A :arr2[16].count,
                                                                                        Tower2B :arr3[16].count,
                                                                                        Tower2C  : arr4[16].count,
                                                                                        Tower2D  : arr5[16].count,
                                                                                        CompletedPercent : ((arr2[16]['count']+arr3[16]['count']+arr4[16]['count']+arr5[16]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                                 
                                                                                    }).then(()=>{
                                                                                        excel.writeRow('StructureActivities',19,{
                                                                                            Activity : arr2[17].milestone,
                                                                                            Tower2A :arr2[17].count,
                                                                                            Tower2B :arr3[17].count,
                                                                                            Tower2C  : arr4[17].count,
                                                                                            Tower2D  : arr5[17].count,
                                                                                            CompletedPercent : ((arr2[17]['count']+arr3[17]['count']+arr4[17]['count']+arr5[17]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                                     
                                                                                        }).then(()=>{
                                                                                            excel.writeRow('StructureActivities',20,{
                                                                                                Activity : arr2[18].milestone,
                                                                                                Tower2A :arr2[18].count,
                                                                                                Tower2B :arr3[18].count,
                                                                                                Tower2C  : arr4[18].count,
                                                                                                Tower2D  : arr5[18].count,
                                                                                                CompletedPercent : ((arr2[18]['count']+arr3[18]['count']+arr4[18]['count']+arr5[18]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                                        
                                                                                            }).then(()=>{
                                                                                                excel.writeRow('StructureActivities',21,{
                                                                                                    Activity : arr2[19].milestone ,
                                                                                                    Tower2A :arr2[19].count,
                                                                                                    Tower2B :arr3[19].count,
                                                                                                    Tower2C  : arr4[19].count,
                                                                                                    Tower2D  : arr5[19].count,
                                                                                                    CompletedPercent : ((arr2[19]['count']+arr3[19]['count']+arr4[19]['count']+arr5[19]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                                            
                                                                                                }).then(()=>{
                                                                                                    excel.writeRow('StructureActivities',22,{
                                                                                                        Activity : arr2[20].milestone,
                                                                                                        Tower2A :arr2[20].count,
                                                                                                        Tower2B :arr3[20].count,
                                                                                                        Tower2C  : arr4[20].count,
                                                                                                        Tower2D  : arr5[20].count,
                                                                                                        CompletedPercent : ((arr2[20]['count']+arr3[20]['count']+arr4[20]['count']+arr5[20]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                                               
                                                                                                    }).then(()=>{
                                                                                                        excel.writeRow('StructureActivities',23,{
                                                                                                            Activity : arr2[21].milestone,
                                                                                                            Tower2A :arr2[21].count,
                                                                                                            Tower2B :arr3[21].count,
                                                                                                            Tower2C  : arr4[21].count,
                                                                                                            Tower2D  : arr5[21].count,
                                                                                                            CompletedPercent : ((arr2[21]['count']+arr3[21]['count']+arr4[21]['count']+arr5[21]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                                                    
                                                                                                        }).then(()=>{
                                                                                                            excel.writeRow('StructureActivities',24,{
                                                                                                                Activity : arr2[22].milestone,
                                                                                                                Tower2A :arr2[22].count,
                                                                                                                Tower2B :arr3[22].count,
                                                                                                                Tower2C  : arr4[22].count,
                                                                                                                Tower2D  : arr5[22].count,
                                                                                                                CompletedPercent : ((arr2[22]['count']+arr3[22]['count']+arr4[22]['count']+arr5[22]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                                                                                                                                                             
                                                                                                            }).then(()=>{
                                                                                                                excel.writeRow('StructureActivities',25,{
                                                                                                                    Activity : arr2[23].milestone,
                                                                                                                    Tower2A :arr2[23].count,
                                                                                                                    Tower2B :arr3[23].count,
                                                                                                                    Tower2C  : arr4[23].count,
                                                                                                                    Tower2D  : arr5[23].count,
                                                                                                                    CompletedPercent : ((arr2[23]['count']+arr3[23]['count']+arr4[23]['count']+arr5[23]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                    
                                                                                                                                             
                                                                                                                }).then(()=>{
                                                                                                                    excel.writeRow('StructureActivities',26,{
                                                                                                                        Activity : arr2[24].milestone,
                                                                                                                        Tower2A :arr2[24].count,
                                                                                                                        Tower2B :arr3[24].count, 
                                                                                                                        Tower2C  : arr4[24].count,
                                                                                                                        Tower2D  : arr5[24].count,
                                                                                                                        CompletedPercent : ((arr2[24]['count']+arr3[24]['count']+arr4[24]['count']+arr5[24]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                                                                                                                                                                                                                   
                                                                                                                    })                                                                                                                        
                                                                                                              });                                                                                                                                                                                                                                           
                                                                                                            });
                                                                                                        });
                                                                                                    });                                                                                                        
                                                                                                });
                                                                                            });
                                                                                        });
                                                                                    });
                                                                                });
                                                                            });
                                                                        });
                                                                    });  
                                                                });
                                                             });
                                                          });
                                                      });
                                                  });
                                              });
                                          });
                                    });
                                });
                            });
                        });
                    });
                });
            });
        return;
}



//function to generate excel greenphase 1
function generateExcelPhase1(arr,arr6,arr7,arr8,arr9,arr10,arr11)
{
    const path = require('path');
    let excel = new Excel(path.join(__dirname,'Greenfieldphase1.xlsx'));
    
    excel.writeSheet('StructureActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','CompletedPercent'],[arr]).then(()=>{
        excel.writeRow('StructureActivities',1,{
            Activity: "Unit Count",
            TowerA : arr[4]['total_unit'],
            TowerB : arr[5]['total_unit'],
            TowerC : arr[6]['total_unit'],
            TowerD : arr[7]['total_unit'],
            TowerE : arr[8]['total_unit'],
            TowerF : arr[9]['total_unit'],
        }).then(()=>{//1 then
                    excel.writeRow('StructureActivities',2,{
                        Activity : arr6[0].milestone,
                        TowerA : arr6[0]['count'],
                        TowerB : arr7[0]['count'],
                        TowerC : arr8[0]['count'],
                        TowerD : arr9[0]['count'],
                        TowerE : arr10[0]['count'],
                        TowerF : arr11[0]['count'],
                        CompletedPercent : ((arr6[0]['count']+arr7[0]['count']+arr8[0]['count']+arr9[0]['count']+arr10[0]['count']+arr11[0]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                        }).then(()=>{//2nd then
                            excel.writeRow('StructureActivities',3,{
                                Activity : arr6[1].milestone,
                                TowerA : arr6[1]['count'],
                                TowerB : arr7[1]['count'],
                                TowerC : arr8[1]['count'],
                                TowerD : arr9[1]['count'],
                                TowerE : arr10[1]['count'],
                                TowerF : arr11[1]['count'],
                                CompletedPercent : ((arr6[1]['count']+arr7[1]['count']+arr8[1]['count']+arr9[1]['count']+arr10[1]['count']+arr11[1]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                }).then(()=>{//3rd then
                                    excel.writeRow('StructureActivities',4,{
                                        Activity : arr6[2].milestone,
                                        TowerA : arr6[2]['count'],
                                        TowerB : arr7[2]['count'],
                                        TowerC : arr8[2]['count'],
                                        TowerD : arr9[2]['count'],
                                        TowerE : arr10[2]['count'],
                                        TowerF : arr11[2]['count'],
                                        CompletedPercent : ((arr6[2]['count']+arr7[2]['count']+arr8[2]['count']+arr9[2]['count']+arr10[2]['count']+arr11[2]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                        }).then(()=>{//4th then
                                            excel.writeRow('StructureActivities',5,{
                                                Activity : arr6[3].milestone,
                                                TowerA : arr6[3]['count'],
                                                TowerB : arr7[3]['count'],
                                                TowerC : arr8[3]['count'],
                                                TowerD : arr9[3]['count'],
                                                TowerE : arr10[3]['count'],
                                                TowerF : arr11[3]['count'],
                                                CompletedPercent : ((arr6[3]['count']+arr7[3]['count']+arr8[3]['count']+arr9[3]['count']+arr10[3]['count']+arr11[3]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                }).then(()=>{//5th then
                                                    excel.writeRow('StructureActivities',6,{
                                                        Activity : arr6[4].milestone,
                                                        TowerA : arr6[4]['count'],
                                                        TowerB : arr7[4]['count'],
                                                        TowerC : arr8[4]['count'],
                                                        TowerD : arr9[4]['count'],
                                                        TowerE : arr10[4]['count'],
                                                        TowerF : arr11[4]['count'],
                                                        CompletedPercent : ((arr6[4]['count']+arr7[4]['count']+arr8[4]['count']+arr9[4]['count']+arr10[4]['count']+arr11[4]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                        }).then(()=>{//6th then
                                                            excel.writeRow('StructureActivities',6,{
                                                                Activity : arr6[4].milestone,
                                                                TowerA : arr6[4]['count'],
                                                                TowerB : arr7[4]['count'],
                                                                TowerC : arr8[4]['count'],
                                                                TowerD : arr9[4]['count'],
                                                                TowerE : arr10[4]['count'],
                                                                TowerF : arr11[4]['count'],
                                                                CompletedPercent : ((arr6[4]['count']+arr7[4]['count']+arr8[4]['count']+arr9[4]['count']+arr10[4]['count']+arr11[4]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                }).then(()=>{//7th then
                                                                    excel.writeRow('StructureActivities',7,{
                                                                        Activity : arr6[4].milestone,
                                                                        TowerA : arr6[4]['count'],
                                                                        TowerB : arr7[4]['count'],
                                                                        TowerC : arr8[4]['count'],
                                                                        TowerD : arr9[4]['count'],
                                                                        TowerE : arr10[4]['count'],
                                                                        TowerF : arr11[4]['count'],
                                                                        CompletedPercent : ((arr6[4]['count']+arr7[4]['count']+arr8[4]['count']+arr9[4]['count']+arr10[4]['count']+arr11[4]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                    }).then(()=>{//8th then
                                                                        excel.writeRow('StructureActivities',8,{
                                                                            Activity : arr6[5].milestone,
                                                                            TowerA : arr6[5]['count'],
                                                                            TowerB : arr7[5]['count'],
                                                                            TowerC : arr8[5]['count'],
                                                                            TowerD : arr9[5]['count'],
                                                                            TowerE : arr10[5]['count'],
                                                                            TowerF : arr11[5]['count'],
                                                                            CompletedPercent : ((arr6[5]['count']+arr7[5]['count']+arr8[5]['count']+arr9[5]['count']+arr10[5]['count']+arr11[5]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                        }).then(()=>{//9th then
                                                                            excel.writeRow('StructureActivities',9,{
                                                                                Activity : arr6[6].milestone,
                                                                                TowerA : arr6[6]['count'],
                                                                                TowerB : arr7[6]['count'],
                                                                                TowerC : arr8[6]['count'],
                                                                                TowerD : arr9[6]['count'],
                                                                                TowerE : arr10[6]['count'],
                                                                                TowerF : arr11[6]['count'],
                                                                                CompletedPercent : ((arr6[6]['count']+arr7[6]['count']+arr8[6]['count']+arr9[6]['count']+arr10[6]['count']+arr11[6]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                            }).then(()=>{//10th then
                                                                                excel.writeRow('StructureActivities',9,{
                                                                                    Activity : arr6[6].milestone,
                                                                                    TowerA : arr6[6]['count'],
                                                                                    TowerB : arr7[6]['count'],
                                                                                    TowerC : arr8[6]['count'],
                                                                                    TowerD : arr9[6]['count'],
                                                                                    TowerE : arr10[6]['count'],
                                                                                    TowerF : arr11[6]['count'],
                                                                                    CompletedPercent : ((arr6[6]['count']+arr7[6]['count']+arr8[6]['count']+arr9[6]['count']+arr10[6]['count']+arr11[6]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                }).then(()=>{//11th then
                                                                                    excel.writeRow('StructureActivities',10,{
                                                                                        Activity : arr6[7].milestone,
                                                                                        TowerA : arr6[7]['count'],
                                                                                        TowerB : arr7[7]['count'],
                                                                                        TowerC : arr8[7]['count'],
                                                                                        TowerD : arr9[7]['count'],
                                                                                        TowerE : arr10[7]['count'],
                                                                                        TowerF : arr11[7]['count'],
                                                                                        CompletedPercent : ((arr6[7]['count']+arr7[7]['count']+arr8[7]['count']+arr9[7]['count']+arr10[7]['count']+arr11[7]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                    }).then(()=>{//12th then
                                                                                        excel.writeRow('StructureActivities',11,{
                                                                                            Activity : arr6[8].milestone,
                                                                                            TowerA : arr6[8]['count'],
                                                                                            TowerB : arr7[8]['count'],
                                                                                            TowerC : arr8[8]['count'],
                                                                                            TowerD : arr9[8]['count'],
                                                                                            TowerE : arr10[8]['count'],
                                                                                            TowerF : arr11[8]['count'],
                                                                                            CompletedPercent : ((arr6[8]['count']+arr7[8]['count']+arr8[8]['count']+arr9[8]['count']+arr10[8]['count']+arr11[8]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                        }).then(()=>{//13th then
                                                                                            excel.writeRow('StructureActivities',12,{
                                                                                                Activity : arr6[9].milestone,
                                                                                                TowerA : arr6[9]['count'],
                                                                                                TowerB : arr7[9]['count'],
                                                                                                TowerC : arr8[9]['count'],
                                                                                                TowerD : arr9[9]['count'],
                                                                                                TowerE : arr10[9]['count'],
                                                                                                TowerF : arr11[9]['count'],
                                                                                                CompletedPercent : ((arr6[9]['count']+arr7[9]['count']+arr8[9]['count']+arr9[9]['count']+arr10[9]['count']+arr11[9]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                            }).then(()=>{//14th then
                                                                                                excel.writeRow('StructureActivities',13,{
                                                                                                    Activity : arr6[10].milestone,
                                                                                                    TowerA : arr6[10]['count'],
                                                                                                    TowerB : arr7[10]['count'],
                                                                                                    TowerC : arr8[10]['count'],
                                                                                                    TowerD : arr9[10]['count'],
                                                                                                    TowerE : arr10[10]['count'],
                                                                                                    TowerF : arr11[10]['count'],
                                                                                                    CompletedPercent : ((arr6[10]['count']+arr7[10]['count']+arr8[10]['count']+arr9[10]['count']+arr10[10]['count']+arr11[10]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                }).then(()=>{//15th then
                                                                                                    excel.writeRow('StructureActivities',14,{
                                                                                                        Activity : arr6[11].milestone,
                                                                                                        TowerA : arr6[11]['count'],
                                                                                                        TowerB : arr7[11]['count'],
                                                                                                        TowerC : arr8[11]['count'],
                                                                                                        TowerD : arr9[11]['count'],
                                                                                                        TowerE : arr10[11]['count'],
                                                                                                        TowerF : arr11[11]['count'],
                                                                                                        CompletedPercent : ((arr6[11]['count']+arr7[11]['count']+arr8[11]['count']+arr9[11]['count']+arr10[11]['count']+arr11[11]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                    }).then(()=>{//16th then
                                                                                                        excel.writeRow('StructureActivities',15,{
                                                                                                            Activity : arr6[12].milestone,
                                                                                                            TowerA : arr6[12]['count'],
                                                                                                            TowerB : arr7[12]['count'],
                                                                                                            TowerC : arr8[12]['count'],
                                                                                                            TowerD : arr9[12]['count'],
                                                                                                            TowerE : arr10[12]['count'],
                                                                                                            TowerF : arr11[12]['count'],
                                                                                                            CompletedPercent : ((arr6[12]['count']+arr7[12]['count']+arr8[12]['count']+arr9[12]['count']+arr10[12]['count']+arr11[12]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                        }).then(()=>{//17th then
                                                                                                            excel.writeRow('StructureActivities',16,{
                                                                                                                Activity : arr6[13].milestone,
                                                                                                                TowerA : arr6[13]['count'],
                                                                                                                TowerB : arr7[13]['count'],
                                                                                                                TowerC : arr8[13]['count'],
                                                                                                                TowerD : arr9[13]['count'],
                                                                                                                TowerE : arr10[13]['count'],
                                                                                                                TowerF : arr11[13]['count'],
                                                                                                                CompletedPercent : ((arr6[13]['count']+arr7[13]['count']+arr8[13]['count']+arr9[13]['count']+arr10[13]['count']+arr11[13]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                            }).then(()=>{//18th then
                                                                                                                excel.writeRow('StructureActivities',17,{
                                                                                                                    Activity : arr6[14].milestone,
                                                                                                                    TowerA : arr6[14]['count'],
                                                                                                                    TowerB : arr7[14]['count'],
                                                                                                                    TowerC : arr8[14]['count'],
                                                                                                                    TowerD : arr9[14]['count'],
                                                                                                                    TowerE : arr10[14]['count'],
                                                                                                                    TowerF : arr11[14]['count'],
                                                                                                                    CompletedPercent : ((arr6[14]['count']+arr7[14]['count']+arr8[14]['count']+arr9[14]['count']+arr10[14]['count']+arr11[14]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                }).then(()=>{ //19th then
                                                                                                                    excel.writeRow('StructureActivities',18,{
                                                                                                                        Activity : arr6[15].milestone,
                                                                                                                        TowerA : arr6[15]['count'],
                                                                                                                        TowerB : arr7[15]['count'],
                                                                                                                        TowerC : arr8[15]['count'],
                                                                                                                        TowerD : arr9[15]['count'],
                                                                                                                        TowerE : arr10[15]['count'],
                                                                                                                        TowerF : arr11[15]['count'],
                                                                                                                        CompletedPercent : ((arr6[15]['count']+arr7[15]['count']+arr8[15]['count']+arr9[15]['count']+arr10[15]['count']+arr11[15]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                    }).then(()=>{//20th then
                                                                                                                        excel.writeRow('StructureActivities',19,{
                                                                                                                            Activity : arr6[16].milestone,
                                                                                                                            TowerA : arr6[16]['count'],
                                                                                                                            TowerB : arr7[16]['count'],
                                                                                                                            TowerC : arr8[16]['count'],
                                                                                                                            TowerD : arr9[16]['count'],
                                                                                                                            TowerE : arr10[16]['count'],
                                                                                                                            TowerF : arr11[16]['count'],
                                                                                                                            CompletedPercent : ((arr6[16]['count']+arr7[16]['count']+arr8[16]['count']+arr9[16]['count']+arr10[16]['count']+arr11[16]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                        }).then(()=>{//21st then
                                                                                                                            excel.writeRow('StructureActivities',20,{
                                                                                                                                Activity : arr6[17].milestone,
                                                                                                                                TowerA : arr6[17]['count'],
                                                                                                                                TowerB : arr7[17]['count'],
                                                                                                                                TowerC : arr8[17]['count'],
                                                                                                                                TowerD : arr9[17]['count'],
                                                                                                                                TowerE : arr10[17]['count'],
                                                                                                                                TowerF : arr11[17]['count'],
                                                                                                                                CompletedPercent : ((arr6[17]['count']+arr7[17]['count']+arr8[17]['count']+arr9[17]['count']+arr10[17]['count']+arr11[17]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                            }).then(()=>{//22nd then
                                                                                                                                excel.writeRow('StructureActivities',21,{
                                                                                                                                    Activity : arr6[18].milestone,
                                                                                                                                    TowerA : arr6[18]['count'],
                                                                                                                                    TowerB : arr7[18]['count'],
                                                                                                                                    TowerC : arr8[18]['count'],
                                                                                                                                    TowerD : arr9[18]['count'],
                                                                                                                                    TowerE : arr10[18]['count'],
                                                                                                                                    TowerF : arr11[18]['count'],
                                                                                                                                    CompletedPercent : ((arr6[18]['count']+arr7[18]['count']+arr8[18]['count']+arr9[18]['count']+arr10[18]['count']+arr11[18]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                }).then(()=>{//23rd then
                                                                                                                                    excel.writeRow('StructureActivities',22,{
                                                                                                                                        Activity : arr6[19].milestone,
                                                                                                                                        TowerA : arr6[19]['count'],
                                                                                                                                        TowerB : arr7[19]['count'],
                                                                                                                                        TowerC : arr8[19]['count'],
                                                                                                                                        TowerD : arr9[19]['count'],
                                                                                                                                        TowerE : arr10[19]['count'],
                                                                                                                                        TowerF : arr11[19]['count'],
                                                                                                                                        CompletedPercent : ((arr6[19]['count']+arr7[19]['count']+arr8[19]['count']+arr9[19]['count']+arr10[19]['count']+arr11[19]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                    }).then(()=>{//24th then
                                                                                                                                        excel.writeRow('StructureActivities',23,{
                                                                                                                                            Activity : arr6[20].milestone,
                                                                                                                                            TowerA : arr6[20]['count'],
                                                                                                                                            TowerB : arr7[20]['count'],
                                                                                                                                            TowerC : arr8[20]['count'],
                                                                                                                                            TowerD : arr9[20]['count'],
                                                                                                                                            TowerE : arr10[20]['count'],
                                                                                                                                            TowerF : arr11[20]['count'],
                                                                                                                                            CompletedPercent : ((arr6[20]['count']+arr7[20]['count']+arr8[20]['count']+arr9[20]['count']+arr10[20]['count']+arr11[20]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                        }).then(()=>{//25th then
                                                                                                                                            excel.writeRow('StructureActivities',24,{
                                                                                                                                                Activity : arr6[21].milestone,
                                                                                                                                                TowerA : arr6[21]['count'],
                                                                                                                                                TowerB : arr7[21]['count'],
                                                                                                                                                TowerC : arr8[21]['count'],
                                                                                                                                                TowerD : arr9[21]['count'],
                                                                                                                                                TowerE : arr10[21]['count'],
                                                                                                                                                TowerF : arr11[21]['count'],
                                                                                                                                                CompletedPercent : ((arr6[21]['count']+arr7[21]['count']+arr8[21]['count']+arr9[21]['count']+arr10[21]['count']+arr11[21]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                            }).then(()=>{//26th then
                                                                                                                                                excel.writeRow('StructureActivities',25,{
                                                                                                                                                    Activity : arr6[22].milestone,
                                                                                                                                                    TowerA : arr6[22]['count'],
                                                                                                                                                    TowerB : arr7[22]['count'],
                                                                                                                                                    TowerC : arr8[22]['count'],
                                                                                                                                                    TowerD : arr9[22]['count'],
                                                                                                                                                    TowerE : arr10[22]['count'],
                                                                                                                                                    TowerF : arr11[22]['count'],
                                                                                                                                                    CompletedPercent : ((arr6[22]['count']+arr7[22]['count']+arr8[22]['count']+arr9[22]['count']+arr10[22]['count']+arr11[22]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                                }).then(()=>{//27th then
                                                                                                                                                    excel.writeRow('StructureActivities',26,{
                                                                                                                                                        Activity : arr6[23].milestone,
                                                                                                                                                        TowerA : arr6[23]['count'],
                                                                                                                                                        TowerB : arr7[23]['count'],
                                                                                                                                                        TowerC : arr8[23]['count'],
                                                                                                                                                        TowerD : arr9[23]['count'],
                                                                                                                                                        TowerE : arr10[23]['count'],
                                                                                                                                                        TowerF : arr11[23]['count'],
                                                                                                                                                        CompletedPercent : ((arr6[23]['count']+arr7[23]['count']+arr8[23]['count']+arr9[23]['count']+arr10[23]['count']+arr11[23]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                                    }).then(()=>{//28th then
                                                                                                                                                        excel.writeRow('StructureActivities',27,{
                                                                                                                                                            Activity : arr6[24].milestone,
                                                                                                                                                            TowerA : arr6[24]['count'],
                                                                                                                                                            TowerB : arr7[24]['count'],
                                                                                                                                                            TowerC : arr8[24]['count'],
                                                                                                                                                            TowerD : arr9[24]['count'],
                                                                                                                                                            TowerE : arr10[24]['count'],
                                                                                                                                                            TowerF : arr11[24]['count'],
                                                                                                                                                            CompletedPercent : ((arr6[24]['count']+arr7[24]['count']+arr8[24]['count']+arr9[24]['count']+arr10[24]['count']+arr11[24]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                                        }).then(()=>{//29th then
                                                                                                                                                            excel.writeRow('StructureActivities',28,{
                                                                                                                                                                Activity : arr6[25].milestone,
                                                                                                                                                                TowerA : arr6[25]['count'],
                                                                                                                                                                TowerB : arr7[25]['count'],
                                                                                                                                                                TowerC : arr8[25]['count'],
                                                                                                                                                                TowerD : arr9[25]['count'],
                                                                                                                                                                TowerE : arr10[25]['count'],
                                                                                                                                                                TowerF : arr11[25]['count'],
                                                                                                                                                                CompletedPercent : ((arr6[25]['count']+arr7[25]['count']+arr8[25]['count']+arr9[25]['count']+arr10[25]['count']+arr11[25]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                                            }).then(()=>{//30th then
                                                                                                                                                                excel.writeRow('StructureActivities',29,{
                                                                                                                                                                    Activity : arr6[26].milestone,
                                                                                                                                                                    TowerA : arr6[26]['count'],
                                                                                                                                                                    TowerB : arr7[26]['count'],
                                                                                                                                                                    TowerC : arr8[26]['count'],
                                                                                                                                                                    TowerD : arr9[26]['count'],
                                                                                                                                                                    TowerE : arr10[26]['count'],
                                                                                                                                                                    TowerF : arr11[26]['count'],
                                                                                                                                                                    CompletedPercent : ((arr6[26]['count']+arr7[26]['count']+arr8[26]['count']+arr9[26]['count']+arr10[26]['count']+arr11[26]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                                                }).then(()=>{//31st then
                                                                                                                                                                    excel.writeRow('StructureActivities',30,{
                                                                                                                                                                        Activity : arr6[27].milestone,
                                                                                                                                                                        TowerA : arr6[27]['count'],
                                                                                                                                                                        TowerB : arr7[27]['count'],
                                                                                                                                                                        TowerC : arr8[27]['count'],
                                                                                                                                                                        TowerD : arr9[27]['count'],
                                                                                                                                                                        TowerE : arr10[27]['count'],
                                                                                                                                                                        TowerF : arr11[27]['count'],
                                                                                                                                                                        CompletedPercent : ((arr6[27]['count']+arr7[27]['count']+arr8[27]['count']+arr9[27]['count']+arr10[27]['count']+arr11[27]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",
                                                                                                                                                                    });
                                                                                                                                                                });//31st then ending                                                                                                                                                                
                                                                                                                                                           });//30th then ending
                                                                                                                                                        });//29th then ending
                                                                                                                                                    });//28th then ending
                                                                                                                                                });//27th then ending
                                                                                                                                            });//26th then ending
                                                                                                                                        });//25th then ending
                                                                                                                                    })//24th then ending
                                                                                                                                })//23rd then ending  
                                                                                                                            });//22nd then ending
                                                                                                                        });//21st then ending
                                                                                                                    });//20th then ending
                                                                                                                });//19th then ending
                                                                                                            });//18th then ending
                                                                                                        });//17th then ending
                                                                                                    })//16th then ending
                                                                                                });//15th then ending
                                                                                            });//14th then ending
                                                                                        });//13th then ending
                                                                                    });//12th then ending
                                                                                });//11th then ending
                                                                            });//10th then ending
                                                                        });//9th then ending
                                                                    });//8th then ending
                                                                });//7th then ending
                                                        });//6th then ending
                                                })//5th then ending
                                        });//4th then ending
                                });//3rd then ending  
                        });//2nd then ending
                    });//1st then ending
    });//excel writesheet
    return;
} 


//function for southren crest
function generateExcelSC(arr,arr12,arr13,arr14)
{

    const path = require('path');
    let excel = new Excel(path.join(__dirname,'SouthrenCrest.xlsx'));
    
    excel.writeSheet('StructureActivities',['Activity','TowerA','TowerB','TowerC','CompletedPercent'],[arr]).then(()=>{
        excel.writeRow('StructureActivities',1,{
            Activity: "Unit Count",
            TowerA : arr[10]['total_unit'],
            TowerB : arr[11]['total_unit'],
            TowerC : arr[12]['total_unit'],
        }).then(()=>{//1 then
                    excel.writeRow('StructureActivities',2,{
                        Activity : arr12[0].milestone,
                        TowerA : arr12[0]['count'],
                        TowerB : arr13[0]['count'],
                        TowerC : arr14[0]['count'],
                        CompletedPercent : ((arr12[0]['count']+arr13[0]['count']+arr14[0]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                        }).then(()=>{//2nd then
                            excel.writeRow('StructureActivities',3,{
                                Activity : arr12[1].milestone,
                                TowerA : arr12[1]['count'],
                                TowerB : arr13[1]['count'],
                                TowerC : arr14[1]['count'],
                                CompletedPercent : ((arr12[1]['count']+arr13[1]['count']+arr14[1]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                }).then(()=>{//3rd then
                                    excel.writeRow('StructureActivities',4,{
                                        Activity : arr12[2].milestone,
                                        TowerA : arr12[2]['count'],
                                        TowerB : arr13[2]['count'],
                                        TowerC : arr14[2]['count'],
                                        CompletedPercent : ((arr12[2]['count']+arr13[2]['count']+arr14[2]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                        }).then(()=>{//4th then
                                            excel.writeRow('StructureActivities',5,{
                                                Activity : arr12[3].milestone,
                                                TowerA : arr12[3]['count'],
                                                TowerB : arr13[3]['count'],
                                                TowerC : arr14[3]['count'],
                                                CompletedPercent : ((arr12[3]['count']+arr13[3]['count']+arr14[3]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                }).then(()=>{//5th then
                                                    excel.writeRow('StructureActivities',6,{
                                                        Activity : arr12[4].milestone,
                                                        TowerA : arr12[4]['count'],
                                                        TowerB : arr13[4]['count'],
                                                        TowerC : arr14[4]['count'],
                                                        CompletedPercent : ((arr12[4]['count']+arr13[4]['count']+arr14[4]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                        }).then(()=>{//6th then
                                                            excel.writeRow('StructureActivities',7,{
                                                                Activity : arr12[4].milestone,
                                                                TowerA : arr12[4]['count'],
                                                                TowerB : arr13[4]['count'],
                                                                TowerC : arr14[4]['count'],
                                                                CompletedPercent : ((arr12[4]['count']+arr13[4]['count']+arr14[4]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                }).then(()=>{//7th then
                                                                    excel.writeRow('StructureActivities',8,{
                                                                        Activity : arr12[4].milestone,
                                                                        TowerA : arr12[4]['count'],
                                                                        TowerB : arr13[4]['count'],
                                                                        TowerC : arr14[4]['count'],
                                                                        CompletedPercent : ((arr12[4]['count']+arr13[4]['count']+arr14[4]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                    }).then(()=>{//8th then
                                                                        excel.writeRow('StructureActivities',9,{
                                                                            Activity : arr12[5].milestone,
                                                                            TowerA : arr12[5]['count'],
                                                                            TowerB : arr13[5]['count'],
                                                                            TowerC : arr14[5]['count'],
                                                                            CompletedPercent : ((arr12[5]['count']+arr13[5]['count']+arr14[5]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                        }).then(()=>{//9th then
                                                                            excel.writeRow('StructureActivities',10,{
                                                                                Activity : arr12[6].milestone,
                                                                                TowerA : arr12[6]['count'],
                                                                                TowerB : arr13[6]['count'],
                                                                                TowerC : arr14[6]['count'],
                                                                                CompletedPercent : ((arr12[6]['count']+arr13[6]['count']+arr14[6]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                            }).then(()=>{//10th then
                                                                                excel.writeRow('StructureActivities',11,{
                                                                                    Activity : arr12[6].milestone,
                                                                                    TowerA : arr12[6]['count'],
                                                                                    TowerB : arr13[6]['count'],
                                                                                    TowerC : arr14[6]['count'],
                                                                                    CompletedPercent : ((arr12[6]['count']+arr13[6]['count']+arr14[6]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                }).then(()=>{//11th then
                                                                                    excel.writeRow('StructureActivities',12,{
                                                                                        Activity : arr12[7].milestone,
                                                                                        TowerA : arr12[7]['count'],
                                                                                        TowerB : arr13[7]['count'],
                                                                                        TowerC : arr14[7]['count'],
                                                                                        CompletedPercent : ((arr12[7]['count']+arr13[7]['count']+arr14[7]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                    }).then(()=>{//12th then
                                                                                        excel.writeRow('StructureActivities',13,{
                                                                                            Activity : arr12[8].milestone,
                                                                                            TowerA : arr12[8]['count'],
                                                                                            TowerB : arr13[8]['count'],
                                                                                            TowerC : arr14[8]['count'],
                                                                                            CompletedPercent : ((arr12[8]['count']+arr13[8]['count']+arr14[8]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                        }).then(()=>{//13th then
                                                                                            excel.writeRow('StructureActivities',14,{
                                                                                                Activity : arr12[9].milestone,
                                                                                                TowerA : arr12[9]['count'],
                                                                                                TowerB : arr13[9]['count'],
                                                                                                TowerC : arr14[9]['count'],
                                                                                                CompletedPercent : ((arr12[9]['count']+arr13[9]['count']+arr14[9]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                            }).then(()=>{//14th then
                                                                                                excel.writeRow('StructureActivities',15,{
                                                                                                    Activity : arr12[10].milestone,
                                                                                                    TowerA : arr12[10]['count'],
                                                                                                    TowerB : arr13[10]['count'],
                                                                                                    TowerC : arr14[10]['count'],
                                                                                                    CompletedPercent : ((arr12[10]['count']+arr13[10]['count']+arr14[10]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                }).then(()=>{//15th then
                                                                                                    excel.writeRow('StructureActivities',16,{
                                                                                                        Activity : arr12[11].milestone,
                                                                                                        TowerA : arr12[11]['count'],
                                                                                                        TowerB : arr13[11]['count'],
                                                                                                        TowerC : arr14[11]['count'],
                                                                                                        CompletedPercent : ((arr12[11]['count']+arr13[11]['count']+arr14[11]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                    }).then(()=>{//16th then
                                                                                                        excel.writeRow('StructureActivities',17,{
                                                                                                            Activity : arr12[12].milestone,
                                                                                                            TowerA : arr12[12]['count'],
                                                                                                            TowerB : arr13[12]['count'],
                                                                                                            TowerC : arr14[12]['count'],
                                                                                                            CompletedPercent : ((arr12[12]['count']+arr13[12]['count']+arr14[12]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                        }).then(()=>{//17th then
                                                                                                            excel.writeRow('StructureActivities',18,{
                                                                                                                Activity : arr12[13].milestone,
                                                                                                                TowerA : arr12[13]['count'],
                                                                                                                TowerB : arr13[13]['count'],
                                                                                                                TowerC : arr14[13]['count'],
                                                                                                                CompletedPercent : ((arr12[13]['count']+arr13[13]['count']+arr14[13]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                            }).then(()=>{//18th then
                                                                                                                excel.writeRow('StructureActivities',19,{
                                                                                                                    Activity : arr12[14].milestone,
                                                                                                                    TowerA : arr12[14]['count'],
                                                                                                                    TowerB : arr13[14]['count'],
                                                                                                                    TowerC : arr14[14]['count'],
                                                                                                                    CompletedPercent : ((arr12[14]['count']+arr13[14]['count']+arr14[14]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                }).then(()=>{ //19th then
                                                                                                                    excel.writeRow('StructureActivities',20,{
                                                                                                                        Activity : arr12[15].milestone,
                                                                                                                        TowerA : arr12[15]['count'],
                                                                                                                        TowerB : arr13[15]['count'],
                                                                                                                        TowerC : arr14[15]['count'],
                                                                                                                        CompletedPercent : ((arr12[15]['count']+arr13[15]['count']+arr14[15]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                    }).then(()=>{//20th then
                                                                                                                        excel.writeRow('StructureActivities',21,{
                                                                                                                            Activity : arr12[16].milestone,
                                                                                                                            TowerA : arr12[16]['count'],
                                                                                                                            TowerB : arr13[16]['count'],
                                                                                                                            TowerC : arr14[16]['count'],
                                                                                                                            CompletedPercent : ((arr12[16]['count']+arr13[16]['count']+arr14[16]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                        }).then(()=>{//21st then
                                                                                                                            excel.writeRow('StructureActivities',22,{
                                                                                                                                Activity : arr12[17].milestone,
                                                                                                                                TowerA : arr12[17]['count'],
                                                                                                                                TowerB : arr13[17]['count'],
                                                                                                                                TowerC : arr14[17]['count'],
                                                                                                                                CompletedPercent : ((arr12[17]['count']+arr13[17]['count']+arr14[17]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                            }).then(()=>{//22nd then
                                                                                                                                excel.writeRow('StructureActivities',23,{
                                                                                                                                    Activity : arr12[18].milestone,
                                                                                                                                    TowerA : arr12[18]['count'],
                                                                                                                                    TowerB : arr13[18]['count'],
                                                                                                                                    TowerC : arr14[18]['count'],
                                                                                                                                    CompletedPercent : ((arr12[18]['count']+arr13[18]['count']+arr14[18]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                }).then(()=>{//23rd then
                                                                                                                                    excel.writeRow('StructureActivities',24,{
                                                                                                                                        Activity : arr12[19].milestone,
                                                                                                                                        TowerA : arr12[19]['count'],
                                                                                                                                        TowerB : arr13[19]['count'],
                                                                                                                                        TowerC : arr14[19]['count'],
                                                                                                                                        CompletedPercent : ((arr12[19]['count']+arr13[19]['count']+arr14[19]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                    }).then(()=>{//24th then
                                                                                                                                        excel.writeRow('StructureActivities',25,{
                                                                                                                                            Activity : arr12[20].milestone,
                                                                                                                                            TowerA : arr12[20]['count'],
                                                                                                                                            TowerB : arr13[20]['count'],
                                                                                                                                            TowerC : arr14[20]['count'],
                                                                                                                                            CompletedPercent : ((arr12[20]['count']+arr13[20]['count']+arr14[20]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                        }).then(()=>{//25th then
                                                                                                                                            excel.writeRow('StructureActivities',26,{
                                                                                                                                                Activity : arr12[21].milestone,
                                                                                                                                                TowerA : arr12[21]['count'],
                                                                                                                                                TowerB : arr13[21]['count'],
                                                                                                                                                TowerC : arr14[21]['count'],
                                                                                                                                                CompletedPercent : ((arr12[21]['count']+arr13[21]['count']+arr14[21]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                            }).then(()=>{//26th then
                                                                                                                                                excel.writeRow('StructureActivities',27,{
                                                                                                                                                    Activity : arr12[22].milestone,
                                                                                                                                                    TowerA : arr12[22]['count'],
                                                                                                                                                    TowerB : arr13[22]['count'],
                                                                                                                                                    TowerC : arr14[22]['count'],
                                                                                                                                                    CompletedPercent : ((arr12[22]['count']+arr13[22]['count']+arr14[22]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                                }).then(()=>{//27th then
                                                                                                                                                    excel.writeRow('StructureActivities',28,{
                                                                                                                                                        Activity : arr12[23].milestone,
                                                                                                                                                        TowerA : arr12[23]['count'],
                                                                                                                                                        TowerB : arr13[23]['count'],
                                                                                                                                                        TowerC : arr14[23]['count'],
                                                                                                                                                        CompletedPercent : ((arr12[23]['count']+arr13[23]['count']+arr14[23]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                                    }).then(()=>{//28th then
                                                                                                                                                        excel.writeRow('StructureActivities',29,{
                                                                                                                                                            Activity : arr12[24].milestone,
                                                                                                                                                            TowerA : arr12[24]['count'],
                                                                                                                                                            TowerB : arr13[24]['count'],
                                                                                                                                                            TowerC : arr14[24]['count'],
                                                                                                                                                            CompletedPercent : ((arr12[24]['count']+arr13[24]['count']+arr14[24]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                                        }).then(()=>{//29th then
                                                                                                                                                            excel.writeRow('StructureActivities',30,{
                                                                                                                                                                Activity : arr12[25].milestone,
                                                                                                                                                                TowerA : arr12[25]['count'],
                                                                                                                                                                TowerB : arr13[25]['count'],
                                                                                                                                                                TowerC : arr14[25]['count'],
                                                                                                                                                                CompletedPercent : ((arr12[25]['count']+arr13[25]['count']+arr14[25]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                                            }).then(()=>{//30th then
                                                                                                                                                                excel.writeRow('StructureActivities',31,{
                                                                                                                                                                    Activity : arr12[26].milestone,
                                                                                                                                                                    TowerA : arr12[26]['count'],
                                                                                                                                                                    TowerB : arr13[26]['count'],
                                                                                                                                                                    TowerC : arr14[26]['count'],
                                                                                                                                                                    CompletedPercent : ((arr12[26]['count']+arr13[26]['count']+arr14[26]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                                                }).then(()=>{//31st then
                                                                                                                                                                    excel.writeRow('StructureActivities',32,{
                                                                                                                                                                        Activity : arr12[27].milestone,
                                                                                                                                                                        TowerA : arr12[27]['count'],
                                                                                                                                                                        TowerB : arr13[27]['count'],
                                                                                                                                                                        TowerC : arr14[27]['count'],
                                                                                                                                                                        CompletedPercent : ((arr12[27]['count']+arr13[27]['count']+arr13[27]['count']+arr14[27]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                                                    }).then(()=>{//32nd then
                                                                                                                                                                        excel.writeRow('StructureActivities',33,{
                                                                                                                                                                            Activity : arr12[28].milestone,
                                                                                                                                                                            TowerA : arr12[28]['count'],
                                                                                                                                                                            TowerB : arr13[28]['count'],
                                                                                                                                                                            TowerC : arr14[28]['count'],
                                                                                                                                                                            CompletedPercent : ((arr12[28]['count']+arr13[28]['count']+arr14[28]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%",
                                                                                                                                                                        })
                                                                                                                                                                    });//32nd then ending
                                                                                                                                                                });//31st then ending                                                                                                                                                                
                                                                                                                                                           });//30th then ending
                                                                                                                                                        });//29th then ending
                                                                                                                                                    });//28th then ending
                                                                                                                                                });//27th then ending
                                                                                                                                            });//26th then ending
                                                                                                                                        });//25th then ending
                                                                                                                                    })//24th then ending
                                                                                                                                })//23rd then ending  
                                                                                                                            });//22nd then ending
                                                                                                                        });//21st then ending
                                                                                                                    });//20th then ending
                                                                                                                });//19th then ending
                                                                                                            });//18th then ending
                                                                                                        });//17th then ending
                                                                                                    })//16th then ending
                                                                                                });//15th then ending
                                                                                            });//14th then ending
                                                                                        });//13th then ending
                                                                                    });//12th then ending
                                                                                });//11th then ending
                                                                            });//10th then ending
                                                                        });//9th then ending
                                                                    });//8th then ending
                                                                });//7th then ending
                                                        });//6th then ending
                                                })//5th then ending
                                        });//4th then ending
                                });//3rd then ending  
                        });//2nd then ending
                    });//1st then ending
    });//excel writesheet
    return;
}

//flat finish excel fror phase 2
function  flatfinishExcelSouthrenCrest(floor_count,arr6,arr7,arr8)
{

            //path to generate excel
            const path = require('path');
            let excel = new Excel(path.join(__dirname,'SouthrenCrest.xlsx'));
    
            excel.writeSheet('FlatFinishActivities',['Activity','TowerA','TowerB','TowerC','CompletedPercent'],[floor_count]).then(()=>{
                excel.writeRow('FlatFinishActivities',1,{
                   Activity: "Floor Count",
                   TowerA : floor_count[0]['total_unit'],
                   TowerB : floor_count[1]['total_unit'],
                   TowerC : floor_count[2]['total_unit'],
               }).then(()=>{//1st then
                   excel.writeRow('FlatFinishActivities',2,{
                    Activity: arr6[0].milestone,
                    TowerA : arr6[0].count,
                    TowerB : arr7[0].count,
                    TowerC : arr8[0].count,
                    CompletedPercent : (((parseInt(arr6[0]['count'],10)+parseInt(arr7[0]['count'],10)+parseInt(arr8[0]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                 }).then(()=>{//2nd then
                    excel.writeRow('FlatFinishActivities',3,{
                        Activity: arr6[1].milestone,
                        TowerA : arr6[1].count,
                        TowerB : arr7[1].count,
                        TowerC : arr8[1].count,
                        CompletedPercent : (((parseInt(arr6[1]['count'],10)+parseInt(arr7[1]['count'],10)+parseInt(arr8[1]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                        }).then(()=>{//3rd then
                        excel.writeRow('FlatFinishActivities',4,{
                            Activity: arr6[2].milestone,
                            TowerA : arr6[2].count,
                            TowerB : arr7[2].count,
                            TowerC : arr8[2].count,
                            CompletedPercent : (((parseInt(arr6[2]['count'],10)+parseInt(arr7[2]['count'],10)+parseInt(arr8[2]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                            }).then(()=>{//4th then
                            excel.writeRow('FlatFinishActivities',5,{
                                Activity: arr6[3].milestone,
                                TowerA : arr6[3].count,
                                TowerB : arr7[3].count,
                                TowerC : arr8[3].count,
                                CompletedPercent : (((parseInt(arr6[3]['count'],10)+parseInt(arr7[3]['count'],10)+parseInt(arr8[3]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                }).then(()=>{//5th then
                                excel.writeRow('FlatFinishActivities',6,{
                                    Activity: arr6[4].milestone,
                                    TowerA : arr6[4].count,
                                    TowerB : arr7[4].count,
                                    TowerC : arr8[4].count,
                                    CompletedPercent : (((parseInt(arr6[4]['count'],10)+parseInt(arr7[4]['count'],10)+parseInt(arr8[4]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                    }).then(()=>{//6th then
                                    excel.writeRow('FlatFinishActivities',7,{
                                        Activity: arr6[5].milestone,
                                        TowerA : arr6[5].count,
                                        TowerB : arr7[5].count,
                                        TowerC : arr8[5].count,
                                        CompletedPercent : (((parseInt(arr6[5]['count'],10)+parseInt(arr7[5]['count'],10)+parseInt(arr8[5]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                        }).then(()=>{//7th then     
                                        excel.writeRow('FlatFinishActivities',8,{
                                            Activity: arr6[6].milestone,
                                            TowerA : arr6[6].count,
                                            TowerB : arr7[6].count,
                                            TowerC : arr8[6].count,
                                            CompletedPercent : (((parseInt(arr6[6]['count'],10)+parseInt(arr7[6]['count'],10)+parseInt(arr8[6]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                            }).then(()=>{//8th then
                                            excel.writeRow('FlatFinishActivities',9,{
                                                Activity: arr6[7].milestone,
                                                TowerA : arr6[7].count,
                                                TowerB : arr7[7].count,
                                                TowerC : arr8[7].count,
                                                CompletedPercent : (((parseInt(arr6[7]['count'],10)+parseInt(arr7[7]['count'],10)+parseInt(arr8[7]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                }).then(()=>{//9th then
                                                excel.writeRow('FlatFinishActivities',10,{
                                                    Activity: arr6[8].milestone,
                                                    TowerA : arr6[8].count,
                                                    TowerB : arr7[8].count,
                                                    TowerC : arr8[8].count,
                                                    CompletedPercent : (((parseInt(arr6[8]['count'],10)+parseInt(arr7[8]['count'],10)+parseInt(arr8[8]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                    }).then(()=>{//10th then
                                                    excel.writeRow('FlatFinishActivities',11,{
                                                        Activity: arr6[9].milestone,
                                                        TowerA : arr6[9].count,
                                                        TowerB : arr7[9].count,
                                                        TowerC : arr8[9].count,
                                                        CompletedPercent : (((parseInt(arr6[9]['count'],10)+parseInt(arr7[9]['count'],10)+parseInt(arr8[9]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                        }).then(()=>{//11th then
                                                        excel.writeRow('FlatFinishActivities',12,{
                                                            Activity: arr6[10].milestone,
                                                            TowerA : arr6[10].count,
                                                            TowerB : arr7[10].count,
                                                            TowerC : arr8[10].count,
                                                            CompletedPercent : (((parseInt(arr6[10]['count'],10)+parseInt(arr7[10]['count'],10)+parseInt(arr8[10]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                            }).then(()=>{//12th then
                                                            excel.writeRow('FlatFinishActivities',13,{
                                                                Activity: arr6[11].milestone,
                                                                TowerA : arr6[11].count,
                                                                TowerB : arr7[11].count,
                                                                TowerC : arr8[11].count,
                                                                CompletedPercent : (((parseInt(arr6[11]['count'],10)+parseInt(arr7[11]['count'],10)+parseInt(arr8[11]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                }).then(()=>{//13th then
                                                                excel.writeRow('FlatFinishActivities',14,{
                                                                    Activity: arr6[12].milestone,
                                                                    TowerA : arr6[12].count,
                                                                    TowerB : arr7[12].count,
                                                                    TowerC : arr8[12].count,
                                                                    CompletedPercent : (((parseInt(arr6[12]['count'],10)+parseInt(arr7[12]['count'],10)+parseInt(arr8[12]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                    }).then(()=>{//14th then
                                                                    excel.writeRow('FlatFinishActivities',15,{
                                                                        Activity: arr6[13].milestone,
                                                                        TowerA : arr6[13].count,
                                                                        TowerB : arr7[13].count,
                                                                        TowerC : arr8[13].count,
                                                                        CompletedPercent : (((parseInt(arr6[13]['count'],10)+parseInt(arr7[13]['count'],10)+parseInt(arr8[13]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                        }).then(()=>{//15th then
                                                                        excel.writeRow('FlatFinishActivities',16,{
                                                                            Activity: arr6[14].milestone,
                                                                            TowerA : arr6[14].count,
                                                                            TowerB : arr7[14].count,
                                                                            TowerC : arr8[14].count,
                                                                            CompletedPercent : (((parseInt(arr6[14]['count'],10)+parseInt(arr7[14]['count'],10)+parseInt(arr8[14]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                            }).then(()=>{//16th
                                                                            excel.writeRow('FlatFinishActivities',17,{
                                                                                Activity: arr6[15].milestone,
                                                                                TowerA : arr6[15].count,
                                                                                TowerB : arr7[15].count,
                                                                                TowerC : arr8[15].count,
                                                                                CompletedPercent : (((parseInt(arr6[15]['count'],10)+parseInt(arr7[15]['count'],10)+parseInt(arr8[15]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                }).then(()=>{//17th
                                                                                excel.writeRow('FlatFinishActivities',18,{
                                                                                    Activity: arr6[16].milestone,
                                                                                    TowerA : arr6[16].count,
                                                                                    TowerB : arr7[16].count,
                                                                                    TowerC : arr8[16].count,
                                                                                    CompletedPercent : (((parseInt(arr6[16]['count'],10)+parseInt(arr7[16]['count'],10)+parseInt(arr8[16]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                    }).then(()=>{//18th
                                                                                    excel.writeRow('FlatFinishActivities',19,{
                                                                                        Activity: arr6[17].milestone,
                                                                                        TowerA : arr6[17].count,
                                                                                        TowerB : arr7[17].count,
                                                                                        TowerC : arr8[17].count,
                                                                                        CompletedPercent : (((parseInt(arr6[17]['count'],10)+parseInt(arr7[17]['count'],10)+parseInt(arr8[17]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                        }).then(()=>{//19th
                                                                                        excel.writeRow('FlatFinishActivities',20,{
                                                                                            Activity: arr6[18].milestone,
                                                                                            TowerA : arr6[18].count,
                                                                                            TowerB : arr7[18].count,
                                                                                            TowerC : arr8[18].count,
                                                                                            CompletedPercent : (((parseInt(arr6[18]['count'],10)+parseInt(arr7[18]['count'],10)+parseInt(arr8[18]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                            }).then(()=>{//20th
                                                                                            excel.writeRow('FlatFinishActivities',21,{
                                                                                                Activity: arr6[19].milestone,
                                                                                                TowerA : arr6[19].count,
                                                                                                TowerB : arr7[19].count,
                                                                                                TowerC : arr8[19].count,
                                                                                                CompletedPercent : (((parseInt(arr6[19]['count'],10)+parseInt(arr7[19]['count'],10)+parseInt(arr8[19]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                }).then(() => {//21st
                                                                                                excel.writeRow('FlatFinishActivities',22,{
                                                                                                    Activity: arr6[20].milestone,
                                                                                                    TowerA : arr6[20].count,
                                                                                                    TowerB : arr7[20].count,
                                                                                                    TowerC : arr8[20].count,
                                                                                                    CompletedPercent : (((parseInt(arr6[20]['count'],10)+parseInt(arr7[20]['count'],10)+parseInt(arr8[20]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                    }).then(() => {//22nd
                                                                                                    excel.writeRow('FlatFinishActivities',23,{
                                                                                                        Activity: arr6[21].milestone,
                                                                                                        TowerA : arr6[21].count,
                                                                                                        TowerB : arr7[21].count,
                                                                                                        TowerC : arr8[21].count,
                                                                                                        CompletedPercent : (((parseInt(arr6[21]['count'],10)+parseInt(arr7[21]['count'],10)+parseInt(arr8[21]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                        }).then(() => {//23rd
                                                                                                        excel.writeRow('FlatFinishActivities',24,{
                                                                                                            Activity: arr6[22].milestone,
                                                                                                            TowerA : arr6[22].count,
                                                                                                            TowerB : arr7[22].count,
                                                                                                            TowerC : arr8[22].count,
                                                                                                            CompletedPercent : (((parseInt(arr6[22]['count'],10)+parseInt(arr7[22]['count'],10)+parseInt(arr8[22]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                            }).then(() => {//24th
                                                                                                            excel.writeRow('FlatFinishActivities',25,{
                                                                                                                Activity: arr6[23].milestone,
                                                                                                                TowerA : arr6[23].count,
                                                                                                                TowerB : arr7[23].count,
                                                                                                                TowerC : arr8[23].count,
                                                                                                                CompletedPercent : (((parseInt(arr6[23]['count'],10)+parseInt(arr7[23]['count'],10)+parseInt(arr8[23]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                }).then(() => {//25th
                                                                                                                excel.writeRow('FlatFinishActivities',26,{
                                                                                                                    Activity: arr6[24].milestone,
                                                                                                                    TowerA : arr6[24].count,
                                                                                                                    TowerB : arr7[24].count,
                                                                                                                    TowerC : arr8[24].count,
                                                                                                                    CompletedPercent : (((parseInt(arr6[24]['count'],10)+parseInt(arr7[24]['count'],10)+parseInt(arr8[24]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                     }).then(() =>{//26th
                                                                                                                    excel.writeRow('FlatFinishActivities',27,{
                                                                                                                        Activity: arr6[25].milestone,
                                                                                                                        TowerA : arr6[25].count,
                                                                                                                        TowerB : arr7[25].count,
                                                                                                                        TowerC : arr8[25].count,
                                                                                                                        CompletedPercent : (((parseInt(arr6[25]['count'],10)+parseInt(arr7[25]['count'],10)+parseInt(arr8[25]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                         }).then(()=>{//27th
                                                                                                                        excel.writeRow('FlatFinishActivities',28,{
                                                                                                                            Activity: arr6[26].milestone,
                                                                                                                            TowerA : arr6[26].count,
                                                                                                                            TowerB : arr7[26].count,
                                                                                                                            TowerC : arr8[26].count,
                                                                                                                            CompletedPercent : (((parseInt(arr6[26]['count'],10)+parseInt(arr7[26]['count'],10)+parseInt(arr8[26]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                        }).then(() =>{//28th
                                                                                                                        excel.writeRow('FlatFinishActivities',29,{
                                                                                                                            Activity: arr6[27].milestone,
                                                                                                                            TowerA : arr6[27].count,
                                                                                                                            TowerB : arr7[27].count,
                                                                                                                            TowerC : arr8[27].count,
                                                                                                                            CompletedPercent : (((parseInt(arr6[27]['count'],10)+parseInt(arr7[27]['count'],10)+parseInt(arr8[27]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                  }).then(()=>{//29th
                                                                                                                    excel.writeRow('FlatFinishActivities',30,{
                                                                                                                        Activity: arr6[28].milestone,
                                                                                                                        TowerA : arr6[28].count,
                                                                                                                        TowerB : arr7[28].count,
                                                                                                                        TowerC : arr8[28].count,
                                                                                                                        CompletedPercent : (((parseInt(arr6[28]['count'],10)+parseInt(arr7[28]['count'],10)+parseInt(arr8[28]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                }).then(()=>{//30th
                                                                                                                        excel.writeRow('FlatFinishActivities',31,{
                                                                                                                            Activity: arr6[29].milestone,
                                                                                                                            TowerA : arr6[29].count,
                                                                                                                            TowerB : arr7[29].count,
                                                                                                                            TowerC : arr8[29].count,
                                                                                                                            CompletedPercent : (((parseInt(arr6[29]['count'],10)+parseInt(arr7[29]['count'],10)+parseInt(arr8[29]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                            }).then(()=>{//31st
                                                                                                                            excel.writeRow('FlatFinishActivities',32,{
                                                                                                                                Activity: arr6[30].milestone,
                                                                                                                                TowerA : arr6[30].count,
                                                                                                                                TowerB : arr7[30].count,
                                                                                                                                TowerC : arr8[30].count,
                                                                                                                                CompletedPercent : (((parseInt(arr6[30]['count'],10)+parseInt(arr7[30]['count'],10)+parseInt(arr8[30]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                }).then(() =>{//32nd
                                                                                                                                excel.writeRow('FlatFinishActivities',33,{
                                                                                                                                    Activity: arr6[31].milestone,
                                                                                                                                    TowerA : arr6[31].count,
                                                                                                                                    TowerB : arr7[31].count,
                                                                                                                                    TowerC : arr8[31].count,
                                                                                                                                    CompletedPercent : (((parseInt(arr6[31]['count'],10)+parseInt(arr7[31]['count'],10)+parseInt(arr8[31]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                    }).then(()=>{//33rd
                                                                                                                                    excel.writeRow('FlatFinishActivities',34,{
                                                                                                                                        Activity: arr6[32].milestone,
                                                                                                                                        TowerA : arr6[32].count,
                                                                                                                                        TowerB : arr7[32].count,
                                                                                                                                        TowerC : arr8[32].count,
                                                                                                                                        CompletedPercent : (((parseInt(arr6[32]['count'],10)+parseInt(arr7[32]['count'],10)+parseInt(arr8[32]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                        }).then(() =>{//34th
                                                                                                                                        excel.writeRow('FlatFinishActivities',35,{
                                                                                                                                            Activity: arr6[33].milestone,
                                                                                                                                            TowerA : arr6[33].count,
                                                                                                                                            TowerB : arr7[33].count,
                                                                                                                                            TowerC : arr8[33].count,
                                                                                                                                            CompletedPercent : (((parseInt(arr6[33]['count'],10)+parseInt(arr7[33]['count'],10)+parseInt(arr8[33]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                            }).then(() => {//35th
                                                                                                                                            excel.writeRow('FlatFinishActivities',36,{
                                                                                                                                                Activity: arr6[34].milestone,
                                                                                                                                                TowerA : arr6[34].count,
                                                                                                                                                TowerB : arr7[34].count,
                                                                                                                                                TowerC : arr8[34].count,
                                                                                                                                                CompletedPercent : (((parseInt(arr6[34]['count'],10)+parseInt(arr7[34]['count'],10)+parseInt(arr8[34]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                               }).then(() => {//36th
                                                                                                                                                excel.writeRow('FlatFinishActivities',37,{
                                                                                                                                                    Activity: arr6[35].milestone,
                                                                                                                                                    TowerA : arr6[35].count,
                                                                                                                                                    TowerB : arr7[35].count,
                                                                                                                                                    TowerC : arr8[35].count,
                                                                                                                                                    CompletedPercent : (((parseInt(arr6[35]['count'],10)+parseInt(arr7[35]['count'],10)+parseInt(arr8[35]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                                    }).then(() => {//37th
                                                                                                                                                        excel.writeRow('FlatFinishActivities',38,{
                                                                                                                                                            Activity: arr6[36].milestone,
                                                                                                                                                            TowerA : arr6[36].count,
                                                                                                                                                            TowerB : arr7[36].count,
                                                                                                                                                            TowerC : arr8[36].count,
                                                                                                                                                            CompletedPercent : (((parseInt(arr6[36]['count'],10)+parseInt(arr7[36]['count'],10)+parseInt(arr8[36]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                                        }).then(() => {//38th
                                                                                                                                                            excel.writeRow('FlatFinishActivities',39,{
                                                                                                                                                                Activity: arr6[37].milestone,
                                                                                                                                                                TowerA : arr6[37].count,
                                                                                                                                                                TowerB : arr7[37].count,
                                                                                                                                                                TowerC : arr8[37].count,
                                                                                                                                                                CompletedPercent : (((parseInt(arr6[37]['count'],10)+parseInt(arr7[37]['count'],10)+parseInt(arr8[37]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                                            })
                                                                                                                                                        });//38th
                                                                                                                                                })//37th
                                                                                                                                            })//36th
                                                                                                                                        })//35th    
                                                                                                                                    })//34th
                                                                                                                                })//33rd
                                                                                                                            })//32nd
                                                                                                                        })//31st
                                                                                                                    })//30th
                                                                                                                  })//29th
                                                                                                                })//28th
                                                                                                              })//27th
                                                                                                            })//26th
                                                                                                          })//25th
                                                                                                        })//24th
                                                                                                    })//23rd 
                                                                                                })//22nd
                                                                                            })//21st
                                                                                        })//20th
                                                                                    })//19th then
                                                                                })//18th then 
                                                                            })//17th then
                                                                        })//16th then
                                                                    })//15th then ending
                                                                })//14th then ending
                                                            })//13th then ending
                                                        })//12th then ending
                                                    })//11th then ending
                                                })//10th then ending
                                            })//9th then ending
                                        })//8th then ending
                                    })//7th then ending
                                });//6th then ending
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;        
}


//flatfinish luxor
function flatfinishExcelLuxor(floor_count,arr6,arr7,arr8,arr9,arr10,arr11,arr12,arr13) 
{
    
            //path to generate excel
            const path = require('path');
            let excel = new Excel(path.join(__dirname,'Luxor.xlsx'));
    
            excel.writeSheet('FlatFinishActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','TowerG','TowerH','CompletedPercent'],[floor_count]).then(()=>{
                excel.writeRow('FlatFinishActivities',1,{
                   Activity: "Floor Count",
                   TowerA : floor_count[0]['total_unit'],
                   TowerB : floor_count[1]['total_unit'],
                   TowerC : floor_count[2]['total_unit'],
                   TowerD : floor_count[3]['total_unit'],
                   TowerE : floor_count[4]['total_unit'],
                   TowerF : floor_count[5]['total_unit'],
                   TowerG : floor_count[6]['total_unit'],
                   TowerH : floor_count[7]['total_unit'],
               }).then(()=>{//1st then
                   excel.writeRow('FlatFinishActivities',2,{
                    Activity: arr6[0].milestone,
                    TowerA : arr6[0].count,
                    TowerB : arr7[0].count,
                    TowerC : arr8[0].count,
                    TowerD : arr9[0].count,
                    TowerE : arr10[0].count,
                    TowerF : arr11[0].count,
                    TowerG : arr12[0].count,
                    TowerH : arr13[0].count,
                    CompletedPercent : (((parseInt(arr6[0]['count'],10)+parseInt(arr7[0]['count'],10)+parseInt(arr8[0]['count'],10)+parseInt(arr9[0]['count'],10)+parseInt(arr10[0]['count'],10)+parseInt(arr11[0]['count'],10)+parseInt(arr12[0]['count'],10)+parseInt(arr13[0]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                 }).then(()=>{//2nd then
                    excel.writeRow('FlatFinishActivities',3,{
                        Activity: arr6[1].milestone,
                        TowerA : arr6[1].count,
                        TowerB : arr7[1].count,
                        TowerC : arr8[1].count,
                        TowerD : arr9[1].count,
                        TowerE : arr10[1].count,
                        TowerF : arr11[1].count,
                        TowerG : arr12[1].count,
                        TowerH : arr13[1].count,    
                        CompletedPercent : (((parseInt(arr6[1]['count'],10)+parseInt(arr7[1]['count'],10)+parseInt(arr8[1]['count'],10)+parseInt(arr9[1]['count'],10)+parseInt(arr10[1]['count'],10)+parseInt(arr11[1]['count'],10)+parseInt(arr12[1]['count'],10)+parseInt(arr13[1]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                        }).then(()=>{//3rd then
                        excel.writeRow('FlatFinishActivities',4,{
                            Activity: arr6[2].milestone,
                            TowerA : arr6[2].count,
                            TowerB : arr7[2].count,
                            TowerC : arr8[2].count,
                            TowerD : arr9[2].count,
                            TowerE : arr10[2].count,
                            TowerF : arr11[2].count,
                            TowerG : arr12[2].count,
                            TowerH : arr13[2].count,        
                            CompletedPercent : (((parseInt(arr6[2]['count'],10)+parseInt(arr7[2]['count'],10)+parseInt(arr8[2]['count'],10)+parseInt(arr9[2]['count'],10)+parseInt(arr10[2]['count'],10)+parseInt(arr11[2]['count'],10)+parseInt(arr12[2]['count'],10)+parseInt(arr13[2]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                            }).then(()=>{//4th then
                            excel.writeRow('FlatFinishActivities',5,{
                                Activity: arr6[3].milestone,
                                TowerA : arr6[3].count,
                                TowerB : arr7[3].count,
                                TowerC : arr8[3].count,
                                TowerD : arr9[3].count,
                                TowerE : arr10[3].count,
                                TowerF : arr11[3].count,
                                TowerG : arr12[3].count,
                                TowerH : arr13[3].count,
                                CompletedPercent : (((parseInt(arr6[3]['count'],10)+parseInt(arr7[3]['count'],10)+parseInt(arr8[3]['count'],10)+parseInt(arr9[3]['count'],10)+parseInt(arr10[3]['count'],10)+parseInt(arr11[3]['count'],10)+parseInt(arr12[3]['count'],10)+parseInt(arr13[3]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                }).then(()=>{//5th then
                                excel.writeRow('FlatFinishActivities',6,{
                                    Activity: arr6[4].milestone,
                                    TowerA : arr6[4].count,
                                    TowerB : arr7[4].count,
                                    TowerC : arr8[4].count,
                                    TowerD : arr9[4].count,
                                    TowerE : arr10[4].count,
                                    TowerF : arr11[4].count,
                                    TowerG : arr12[4].count,
                                    TowerH : arr13[4].count,    
                                    CompletedPercent : (((parseInt(arr6[4]['count'],10)+parseInt(arr7[4]['count'],10)+parseInt(arr8[4]['count'],10)+parseInt(arr9[4]['count'],10)+parseInt(arr10[4]['count'],10)+parseInt(arr11[4]['count'],10)+parseInt(arr12[4]['count'],10)+parseInt(arr13[4]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                    }).then(()=>{//6th then
                                    excel.writeRow('FlatFinishActivities',7,{
                                        Activity: arr6[5].milestone,
                                        TowerA : arr6[5].count,
                                        TowerB : arr7[5].count,
                                        TowerC : arr8[5].count,
                                        TowerD : arr9[5].count,
                                        TowerE : arr10[5].count,
                                        TowerF : arr11[5].count,
                                        TowerG : arr12[5].count,
                                        TowerH : arr13[5].count,        
                                        CompletedPercent : (((parseInt(arr6[5]['count'],10)+parseInt(arr7[5]['count'],10)+parseInt(arr8[5]['count'],10)+parseInt(arr9[5]['count'],10)+parseInt(arr10[5]['count'],10)+parseInt(arr11[5]['count'],10)+parseInt(arr12[5]['count'],10)+parseInt(arr13[5]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                        }).then(()=>{//7th then     
                                        excel.writeRow('FlatFinishActivities',8,{
                                            Activity: arr6[6].milestone,
                                            TowerA : arr6[6].count,
                                            TowerB : arr7[6].count,
                                            TowerC : arr8[6].count,
                                            TowerD : arr9[6].count,
                                            TowerE : arr10[6].count,
                                            TowerF : arr11[6].count,
                                            TowerG : arr12[6].count,
                                            TowerH : arr13[6].count,            
                                            CompletedPercent : (((parseInt(arr6[6]['count'],10)+parseInt(arr7[6]['count'],10)+parseInt(arr8[6]['count'],10)+parseInt(arr9[6]['count'],10)+parseInt(arr10[6]['count'],10)+parseInt(arr11[6]['count'],10)+parseInt(arr12[6]['count'],10)+parseInt(arr13[6]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                            }).then(()=>{//8th then
                                            excel.writeRow('FlatFinishActivities',9,{
                                                Activity: arr6[7].milestone,
                                                TowerA : arr6[7].count,
                                                TowerB : arr7[7].count,
                                                TowerC : arr8[7].count,
                                                TowerD : arr9[7].count,
                                                TowerE : arr10[7].count,
                                                TowerF : arr11[7].count,
                                                TowerG : arr12[7].count,
                                                TowerH : arr13[7].count,                
                                                CompletedPercent : (((parseInt(arr6[7]['count'],10)+parseInt(arr7[7]['count'],10)+parseInt(arr8[7]['count'],10)+parseInt(arr9[7]['count'],10)+parseInt(arr10[7]['count'],10)+parseInt(arr11[7]['count'],10)+parseInt(arr12[7]['count'],10)+parseInt(arr13[7]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                }).then(()=>{//9th then
                                                excel.writeRow('FlatFinishActivities',10,{
                                                    Activity: arr6[8].milestone,
                                                    TowerA : arr6[8].count,
                                                    TowerB : arr7[8].count,
                                                    TowerC : arr8[8].count,
                                                    TowerD : arr9[8].count,
                                                    TowerE : arr10[8].count,
                                                    TowerF : arr11[8].count,
                                                    TowerG : arr12[8].count,
                                                    TowerH : arr13[8].count,                    
                                                    CompletedPercent : (((parseInt(arr6[8]['count'],10)+parseInt(arr7[8]['count'],10)+parseInt(arr8[8]['count'],10)+parseInt(arr9[8]['count'],10)+parseInt(arr10[8]['count'],10)+parseInt(arr11[8]['count'],10)+parseInt(arr12[8]['count'],10)+parseInt(arr13[8]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                    }).then(()=>{//10th then
                                                    excel.writeRow('FlatFinishActivities',11,{
                                                        Activity: arr6[9].milestone,
                                                        TowerA : arr6[9].count,
                                                        TowerB : arr7[9].count,
                                                        TowerC : arr8[9].count,
                                                        TowerD : arr9[9].count,
                                                        TowerE : arr10[9].count,
                                                        TowerF : arr11[9].count,
                                                        TowerG : arr12[9].count,
                                                        TowerH : arr13[9].count,                        
                                                        CompletedPercent : (((parseInt(arr6[9]['count'],10)+parseInt(arr7[9]['count'],10)+parseInt(arr8[9]['count'],10)+parseInt(arr9[9]['count'],10)+parseInt(arr10[9]['count'],10)+parseInt(arr11[9]['count'],10)+parseInt(arr12[9]['count'],10)+parseInt(arr13[9]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                        }).then(()=>{//11th then
                                                        excel.writeRow('FlatFinishActivities',12,{
                                                            Activity: arr6[10].milestone,
                                                            TowerA : arr6[10].count,
                                                            TowerB : arr7[10].count,
                                                            TowerC : arr8[10].count,
                                                            TowerD : arr9[10].count,
                                                            TowerE : arr10[10].count,
                                                            TowerF : arr11[10].count,
                                                            TowerG : arr12[10].count,
                                                            TowerH : arr13[10].count,                            
                                                            CompletedPercent : (((parseInt(arr6[10]['count'],10)+parseInt(arr7[10]['count'],10)+parseInt(arr8[10]['count'],10)+parseInt(arr9[10]['count'],10)+parseInt(arr10[10]['count'],10)+parseInt(arr11[10]['count'],10)+parseInt(arr12[10]['count'],10)+parseInt(arr13[10]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                            }).then(()=>{//12th then
                                                            excel.writeRow('FlatFinishActivities',13,{
                                                                Activity: arr6[11].milestone,
                                                                TowerA : arr6[11].count,
                                                                TowerB : arr7[11].count,
                                                                TowerC : arr8[11].count,
                                                                TowerD : arr9[11].count,
                                                                TowerE : arr10[11].count,
                                                                TowerF : arr11[11].count,
                                                                TowerG : arr12[11].count,
                                                                TowerH : arr13[11].count,    
                                                                CompletedPercent : (((parseInt(arr6[11]['count'],10)+parseInt(arr7[11]['count'],10)+parseInt(arr8[11]['count'],10)+parseInt(arr9[11]['count'],10)+parseInt(arr10[11]['count'],10)+parseInt(arr11[11]['count'],10)+parseInt(arr12[11]['count'],10)+parseInt(arr13[11]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                }).then(()=>{//13th then
                                                                excel.writeRow('FlatFinishActivities',14,{
                                                                    Activity: arr6[12].milestone,
                                                                    TowerA : arr6[12].count,
                                                                    TowerB : arr7[12].count,
                                                                    TowerC : arr8[12].count,
                                                                    TowerD : arr9[12].count,
                                                                    TowerE : arr10[12].count,
                                                                    TowerF : arr11[12].count,
                                                                    TowerG : arr12[12].count,
                                                                    TowerH : arr13[12].count,                                    
                                                                    CompletedPercent : (((parseInt(arr6[12]['count'],10)+parseInt(arr7[12]['count'],10)+parseInt(arr8[12]['count'],10)+parseInt(arr9[12]['count'],10)+parseInt(arr10[12]['count'],10)+parseInt(arr11[12]['count'],10)+parseInt(arr12[12]['count'],10)+parseInt(arr13[12]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                    }).then(()=>{//14th then
                                                                    excel.writeRow('FlatFinishActivities',15,{
                                                                        Activity: arr6[13].milestone,
                                                                        TowerA : arr6[13].count,
                                                                        TowerB : arr7[13].count,
                                                                        TowerC : arr8[13].count,
                                                                        TowerD : arr9[13].count,
                                                                        TowerE : arr10[13].count,
                                                                        TowerF : arr11[13].count,
                                                                        TowerG : arr12[13].count,
                                                                        TowerH : arr13[13].count,                                        
                                                                        CompletedPercent : (((parseInt(arr6[13]['count'],10)+parseInt(arr7[13]['count'],10)+parseInt(arr8[13]['count'],10)+parseInt(arr9[13]['count'],10)+parseInt(arr10[13]['count'],10)+parseInt(arr11[13]['count'],10)+parseInt(arr12[13]['count'],10)+parseInt(arr13[13]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                        }).then(()=>{//15th then
                                                                        excel.writeRow('FlatFinishActivities',16,{
                                                                            Activity: arr6[14].milestone,
                                                                            TowerA : arr6[14].count,
                                                                            TowerB : arr7[14].count,
                                                                            TowerC : arr8[14].count,
                                                                            TowerD : arr9[14].count,
                                                                            TowerE : arr10[14].count,
                                                                            TowerF : arr11[14].count,
                                                                            TowerG : arr12[14].count,
                                                                            TowerH : arr13[14].count,                                            
                                                                            CompletedPercent : (((parseInt(arr6[14]['count'],10)+parseInt(arr7[14]['count'],10)+parseInt(arr8[14]['count'],10)+parseInt(arr9[14]['count'],10)+parseInt(arr10[14]['count'],10)+parseInt(arr11[14]['count'],10)+parseInt(arr12[14]['count'],10)+parseInt(arr13[14]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                            }).then(()=>{//16th
                                                                            excel.writeRow('FlatFinishActivities',17,{
                                                                                Activity: arr6[15].milestone,
                                                                                TowerA : arr6[15].count,
                                                                                TowerB : arr7[15].count,
                                                                                TowerC : arr8[15].count,
                                                                                TowerD : arr9[15].count,
                                                                                TowerE : arr10[15].count,
                                                                                TowerF : arr11[15].count,
                                                                                TowerG : arr12[15].count,
                                                                                TowerH : arr13[15].count,                                                
                                                                                CompletedPercent : (((parseInt(arr6[15]['count'],10)+parseInt(arr7[15]['count'],10)+parseInt(arr8[15]['count'],10)+parseInt(arr9[15]['count'],10)+parseInt(arr10[15]['count'],10)+parseInt(arr11[15]['count'],10)+parseInt(arr12[15]['count'],10)+parseInt(arr13[15]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                }).then(()=>{//17th
                                                                                excel.writeRow('FlatFinishActivities',18,{
                                                                                    Activity: arr6[16].milestone,
                                                                                    TowerA : arr6[16].count,
                                                                                    TowerB : arr7[16].count,
                                                                                    TowerC : arr8[16].count,
                                                                                    TowerD : arr9[16].count,
                                                                                    TowerE : arr10[16].count,
                                                                                    TowerF : arr11[16].count,
                                                                                    TowerG : arr12[16].count,
                                                                                    TowerH : arr13[16].count,                                                    
                                                                                    CompletedPercent : (((parseInt(arr6[16]['count'],10)+parseInt(arr7[16]['count'],10)+parseInt(arr8[16]['count'],10)+parseInt(arr9[16]['count'],10)+parseInt(arr10[16]['count'],10)+parseInt(arr11[16]['count'],10)+parseInt(arr12[16]['count'],10)+parseInt(arr13[16]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                    }).then(()=>{//18th
                                                                                    excel.writeRow('FlatFinishActivities',19,{
                                                                                        Activity: arr6[17].milestone,
                                                                                        TowerA : arr6[17].count,
                                                                                        TowerB : arr7[17].count,
                                                                                        TowerC : arr8[17].count,
                                                                                        TowerD : arr9[17].count,
                                                                                        TowerE : arr10[17].count,
                                                                                        TowerF : arr11[17].count,
                                                                                        TowerG : arr12[17].count,
                                                                                        TowerH : arr13[17].count,                                                        
                                                                                        CompletedPercent : (((parseInt(arr6[17]['count'],10)+parseInt(arr7[17]['count'],10)+parseInt(arr8[17]['count'],10)+parseInt(arr9[17]['count'],10)+parseInt(arr10[17]['count'],10)+parseInt(arr11[17]['count'],10)+parseInt(arr12[17]['count'],10)+parseInt(arr13[17]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                        }).then(()=>{//19th
                                                                                        excel.writeRow('FlatFinishActivities',20,{
                                                                                            Activity: arr6[18].milestone,
                                                                                            TowerA : arr6[18].count,
                                                                                            TowerB : arr7[18].count,
                                                                                            TowerC : arr8[18].count,
                                                                                            TowerD : arr9[18].count,
                                                                                            TowerE : arr10[18].count,
                                                                                            TowerF : arr11[18].count,
                                                                                            TowerG : arr12[18].count,
                                                                                            TowerH : arr13[18].count,                                                            
                                                                                            CompletedPercent : (((parseInt(arr6[18]['count'],10)+parseInt(arr7[18]['count'],10)+parseInt(arr8[18]['count'],10)+parseInt(arr9[18]['count'],10)+parseInt(arr10[18]['count'],10)+parseInt(arr11[18]['count'],10)+parseInt(arr12[18]['count'],10)+parseInt(arr13[18]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                            }).then(()=>{//20th
                                                                                            excel.writeRow('FlatFinishActivities',21,{
                                                                                                Activity: arr6[19].milestone,
                                                                                                TowerA : arr6[19].count,
                                                                                                TowerB : arr7[19].count,
                                                                                                TowerC : arr8[19].count,
                                                                                                TowerD : arr9[19].count,
                                                                                                TowerE : arr10[19].count,
                                                                                                TowerF : arr11[19].count,
                                                                                                TowerG : arr12[19].count,
                                                                                                TowerH : arr13[19].count,                                                                
                                                                                                CompletedPercent : (((parseInt(arr6[19]['count'],10)+parseInt(arr7[19]['count'],10)+parseInt(arr8[19]['count'],10)+parseInt(arr9[19]['count'],10)+parseInt(arr10[19]['count'],10)+parseInt(arr11[19]['count'],10)+parseInt(arr12[19]['count'],10)+parseInt(arr13[19]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                }).then(() => {//21st
                                                                                                excel.writeRow('FlatFinishActivities',22,{
                                                                                                    Activity: arr6[20].milestone,
                                                                                                    TowerA : arr6[20].count,
                                                                                                    TowerB : arr7[20].count,
                                                                                                    TowerC : arr8[20].count,
                                                                                                    TowerD : arr9[20].count,
                                                                                                    TowerE : arr10[20].count,
                                                                                                    TowerF : arr11[20].count,
                                                                                                    TowerG : arr12[20].count,
                                                                                                    TowerH : arr13[20].count,                                                                    
                                                                                                    CompletedPercent : (((parseInt(arr6[20]['count'],10)+parseInt(arr7[20]['count'],10)+parseInt(arr8[20]['count'],10)+parseInt(arr9[20]['count'],10)+parseInt(arr10[20]['count'],10)+parseInt(arr11[20]['count'],10)+parseInt(arr12[20]['count'],10)+parseInt(arr13[20]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                    }).then(() => {//22nd
                                                                                                    excel.writeRow('FlatFinishActivities',23,{
                                                                                                        Activity: arr6[21].milestone,
                                                                                                        TowerA : arr6[21].count,
                                                                                                        TowerB : arr7[21].count,
                                                                                                        TowerC : arr8[21].count,
                                                                                                        TowerD : arr9[21].count,
                                                                                                        TowerE : arr10[21].count,
                                                                                                        TowerF : arr11[21].count,
                                                                                                        TowerG : arr12[21].count,
                                                                                                        TowerH : arr13[21].count,    
                                                                                                        CompletedPercent : (((parseInt(arr6[21]['count'],10)+parseInt(arr7[21]['count'],10)+parseInt(arr8[21]['count'],10)+parseInt(arr9[21]['count'],10)+parseInt(arr10[21]['count'],10)+parseInt(arr11[21]['count'],10)+parseInt(arr12[21]['count'],10)+parseInt(arr13[21]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                        }).then(() => {//23rd
                                                                                                        excel.writeRow('FlatFinishActivities',24,{
                                                                                                            Activity: arr6[22].milestone,
                                                                                                            TowerA : arr6[22].count,
                                                                                                            TowerB : arr7[22].count,
                                                                                                            TowerC : arr8[22].count,
                                                                                                            TowerD : arr9[22].count,
                                                                                                            TowerE : arr10[22].count,
                                                                                                            TowerF : arr11[22].count,
                                                                                                            TowerG : arr12[22].count,
                                                                                                            TowerH : arr13[22].count,                                                                            
                                                                                                            CompletedPercent : (((parseInt(arr6[22]['count'],10)+parseInt(arr7[22]['count'],10)+parseInt(arr8[22]['count'],10)+parseInt(arr9[22]['count'],10)+parseInt(arr10[22]['count'],10)+parseInt(arr11[22]['count'],10)+parseInt(arr12[22]['count'],10)+parseInt(arr13[22]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                            }).then(() => {//24th
                                                                                                            excel.writeRow('FlatFinishActivities',25,{
                                                                                                                Activity: arr6[23].milestone,
                                                                                                                TowerA : arr6[23].count,
                                                                                                                TowerB : arr7[23].count,
                                                                                                                TowerC : arr8[23].count,
                                                                                                                TowerD : arr9[23].count,
                                                                                                                TowerE : arr10[23].count,
                                                                                                                TowerF : arr11[23].count,
                                                                                                                TowerG : arr12[23].count,
                                                                                                                TowerH : arr13[23].count,                                                                                
                                                                                                                CompletedPercent : (((parseInt(arr6[23]['count'],10)+parseInt(arr7[23]['count'],10)+parseInt(arr8[23]['count'],10)+parseInt(arr9[23]['count'],10)+parseInt(arr10[23]['count'],10)+parseInt(arr11[23]['count'],10)+parseInt(arr12[23]['count'],10)+parseInt(arr13[23]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                }).then(() => {//25th
                                                                                                                excel.writeRow('FlatFinishActivities',26,{
                                                                                                                    Activity: arr6[24].milestone,
                                                                                                                    TowerA : arr6[24].count,
                                                                                                                    TowerB : arr7[24].count,
                                                                                                                    TowerC : arr8[24].count,
                                                                                                                    TowerD : arr9[24].count,
                                                                                                                    TowerE : arr10[24].count,
                                                                                                                    TowerF : arr11[24].count,
                                                                                                                    TowerG : arr12[24].count,
                                                                                                                    TowerH : arr13[24].count,                                                                                    
                                                                                                                    CompletedPercent : (((parseInt(arr6[24]['count'],10)+parseInt(arr7[24]['count'],10)+parseInt(arr8[24]['count'],10)+parseInt(arr9[24]['count'],10)+parseInt(arr10[24]['count'],10)+parseInt(arr11[24]['count'],10)+parseInt(arr12[24]['count'],10)+parseInt(arr13[24]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                     }).then(() =>{//26th
                                                                                                                    excel.writeRow('FlatFinishActivities',27,{
                                                                                                                        Activity: arr6[25].milestone,
                                                                                                                        TowerA : arr6[25].count,
                                                                                                                        TowerB : arr7[25].count,
                                                                                                                        TowerC : arr8[25].count,
                                                                                                                        TowerD : arr9[25].count,
                                                                                                                        TowerE : arr10[25].count,
                                                                                                                        TowerF : arr11[25].count,
                                                                                                                        TowerG : arr12[25].count,
                                                                                                                        TowerH : arr13[25].count,                                                                                        
                                                                                                                        CompletedPercent : (((parseInt(arr6[25]['count'],10)+parseInt(arr7[25]['count'],10)+parseInt(arr8[25]['count'],10)+parseInt(arr9[25]['count'],10)+parseInt(arr10[25]['count'],10)+parseInt(arr11[25]['count'],10)+parseInt(arr12[25]['count'],10)+parseInt(arr13[25]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                         }).then(()=>{//27th
                                                                                                                        excel.writeRow('FlatFinishActivities',28,{
                                                                                                                            Activity: arr6[26].milestone,
                                                                                                                            TowerA : arr6[26].count,
                                                                                                                            TowerB : arr7[26].count,
                                                                                                                            TowerC : arr8[26].count,
                                                                                                                            TowerD : arr9[26].count,
                                                                                                                            TowerE : arr10[26].count,
                                                                                                                            TowerF : arr11[26].count,
                                                                                                                            TowerG : arr12[26].count,
                                                                                                                            TowerH : arr13[26].count,                                                                                            
                                                                                                                            CompletedPercent : (((parseInt(arr6[26]['count'],10)+parseInt(arr7[26]['count'],10)+parseInt(arr8[26]['count'],10)+parseInt(arr9[26]['count'],10)+parseInt(arr10[26]['count'],10)+parseInt(arr11[26]['count'],10)+parseInt(arr12[26]['count'],10)+parseInt(arr13[26]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                        }).then(() =>{//28th
                                                                                                                        excel.writeRow('FlatFinishActivities',29,{
                                                                                                                            Activity: arr6[27].milestone,
                                                                                                                            TowerA : arr6[27].count,
                                                                                                                            TowerB : arr7[27].count,
                                                                                                                            TowerC : arr8[27].count,
                                                                                                                            TowerD : arr9[27].count,
                                                                                                                            TowerE : arr10[27].count,
                                                                                                                            TowerF : arr11[27].count,
                                                                                                                            TowerG : arr12[27].count,
                                                                                                                            TowerH : arr13[27].count,    
                                                                                                                            CompletedPercent : (((parseInt(arr6[27]['count'],10)+parseInt(arr7[27]['count'],10)+parseInt(arr8[27]['count'],10)+parseInt(arr9[27]['count'],10)+parseInt(arr10[27]['count'],10)+parseInt(arr11[27]['count'],10)+parseInt(arr12[27]['count'],10)+parseInt(arr13[27]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                  }).then(()=>{//29th
                                                                                                                    excel.writeRow('FlatFinishActivities',30,{
                                                                                                                        Activity: arr6[28].milestone,
                                                                                                                        TowerA : arr6[28].count,
                                                                                                                        TowerB : arr7[28].count,
                                                                                                                        TowerC : arr8[28].count,
                                                                                                                        TowerD : arr9[28].count,
                                                                                                                        TowerE : arr10[28].count,
                                                                                                                        TowerF : arr11[28].count,
                                                                                                                        TowerG : arr12[28].count,
                                                                                                                        TowerH : arr13[28].count,                                                                                        
                                                                                                                        CompletedPercent : (((parseInt(arr6[28]['count'],10)+parseInt(arr7[28]['count'],10)+parseInt(arr8[28]['count'],10)+parseInt(arr9[28]['count'],10)+parseInt(arr10[28]['count'],10)+parseInt(arr11[28]['count'],10)+parseInt(arr12[28]['count'],10)+parseInt(arr13[28]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                }).then(()=>{//30th
                                                                                                                        excel.writeRow('FlatFinishActivities',31,{
                                                                                                                            Activity: arr6[29].milestone,
                                                                                                                            TowerA : arr6[29].count,
                                                                                                                            TowerB : arr7[29].count,
                                                                                                                            TowerC : arr8[29].count,
                                                                                                                            TowerD : arr9[29].count,
                                                                                                                            TowerE : arr10[29].count,
                                                                                                                            TowerF : arr11[29].count,
                                                                                                                            TowerG : arr12[29].count,
                                                                                                                            TowerH : arr13[29].count,                                                                                            
                                                                                                                            CompletedPercent : (((parseInt(arr6[29]['count'],10)+parseInt(arr7[29]['count'],10)+parseInt(arr8[29]['count'],10)+parseInt(arr9[29]['count'],10)+parseInt(arr10[29]['count'],10)+parseInt(arr11[29]['count'],10)+parseInt(arr12[29]['count'],10)+parseInt(arr13[29]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                            }).then(()=>{//31st
                                                                                                                            excel.writeRow('FlatFinishActivities',32,{
                                                                                                                                Activity: arr6[30].milestone,
                                                                                                                                TowerA : arr6[30].count,
                                                                                                                                TowerB : arr7[30].count,
                                                                                                                                TowerC : arr8[30].count,
                                                                                                                                TowerD : arr9[30].count,
                                                                                                                                TowerE : arr10[30].count,
                                                                                                                                TowerF : arr11[30].count,
                                                                                                                                TowerG : arr12[30].count,
                                                                                                                                TowerH : arr13[30].count,                                                                                                
                                                                                                                                CompletedPercent : (((parseInt(arr6[30]['count'],10)+parseInt(arr7[30]['count'],10)+parseInt(arr8[30]['count'],10)+parseInt(arr9[30]['count'],10)+parseInt(arr10[30]['count'],10)+parseInt(arr11[30]['count'],10)+parseInt(arr12[30]['count'],10)+parseInt(arr13[30]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                }).then(() =>{//32nd
                                                                                                                                excel.writeRow('FlatFinishActivities',33,{
                                                                                                                                    Activity: arr6[31].milestone,
                                                                                                                                    TowerA : arr6[31].count,
                                                                                                                                    TowerB : arr7[31].count,
                                                                                                                                    TowerC : arr8[31].count,
                                                                                                                                    TowerD : arr9[31].count,
                                                                                                                                    TowerE : arr10[31].count,
                                                                                                                                    TowerF : arr11[31].count,
                                                                                                                                    TowerG : arr12[31].count,
                                                                                                                                    TowerH : arr13[31].count,                                                                                                    
                                                                                                                                    CompletedPercent : (((parseInt(arr6[31]['count'],10)+parseInt(arr7[31]['count'],10)+parseInt(arr8[31]['count'],10)+parseInt(arr9[31]['count'],10)+parseInt(arr10[31]['count'],10)+parseInt(arr11[31]['count'],10)+parseInt(arr12[31]['count'],10)+parseInt(arr13[31]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                    }).then(()=>{//33rd
                                                                                                                                    excel.writeRow('FlatFinishActivities',34,{
                                                                                                                                        Activity: arr6[32].milestone,
                                                                                                                                        TowerA : arr6[32].count,
                                                                                                                                        TowerB : arr7[32].count,
                                                                                                                                        TowerC : arr8[32].count,
                                                                                                                                        TowerD : arr9[32].count,                                    
                                                                                                                                        TowerE : arr10[32].count,
                                                                                                                                        TowerF : arr11[32].count,
                                                                                                                                        TowerG : arr12[32].count,
                                                                                                                                        TowerH : arr13[32].count,                                                                                                        
                                                                                                                                        CompletedPercent : (((parseInt(arr6[32]['count'],10)+parseInt(arr7[32]['count'],10)+parseInt(arr8[32]['count'],10)+parseInt(arr9[32]['count'],10)+parseInt(arr10[32]['count'],10)+parseInt(arr11[32]['count'],10)+parseInt(arr12[32]['count'],10)+parseInt(arr13[32]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                        }).then(() =>{//34th
                                                                                                                                        excel.writeRow('FlatFinishActivities',35,{
                                                                                                                                            Activity: arr6[33].milestone,
                                                                                                                                            TowerA : arr6[33].count,
                                                                                                                                            TowerB : arr7[33].count,
                                                                                                                                            TowerC : arr8[33].count,
                                                                                                                                            TowerD : arr9[33].count,
                                                                                                                                            TowerE : arr10[33].count,
                                                                                                                                            TowerF : arr11[33].count,
                                                                                                                                            TowerG : arr12[33].count,
                                                                                                                                            TowerH : arr13[33].count,    
                                                                                                        
                                                                                                                                            CompletedPercent : (((parseInt(arr6[33]['count'],10)+parseInt(arr7[33]['count'],10)+parseInt(arr8[33]['count'],10)+parseInt(arr9[33]['count'],10)+parseInt(arr10[33]['count'],10)+parseInt(arr11[33]['count'],10)+parseInt(arr12[33]['count'],10)+parseInt(arr13[33]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                            }).then(() => {//35th
                                                                                                                                            excel.writeRow('FlatFinishActivities',36,{
                                                                                                                                                Activity: arr6[34].milestone,
                                                                                                                                                TowerA : arr6[34].count,
                                                                                                                                                TowerB : arr7[34].count,
                                                                                                                                                TowerC : arr8[34].count,
                                                                                                                                                TowerD : arr9[34].count,
                                                                                                                                                TowerE : arr10[34].count,
                                                                                                                                                TowerF : arr11[34].count,
                                                                                                                                                TowerG : arr12[34].count,
                                                                                                                                                TowerH : arr13[34].count,    
                                                                                                            
                                                                                                                                                CompletedPercent : (((parseInt(arr6[34]['count'],10)+parseInt(arr7[34]['count'],10)+parseInt(arr8[34]['count'],10)+parseInt(arr9[34]['count'],10)+parseInt(arr10[34]['count'],10)+parseInt(arr11[34]['count'],10)+parseInt(arr12[34]['count'],10)+parseInt(arr13[34]['count'],10))/(floor_count[0]['total_unit']+floor_count[1]['total_unit']+floor_count[2]['total_unit']+floor_count[3]['total_unit']+floor_count[4]['total_unit']+floor_count[5]['total_unit']+floor_count[6]['total_unit']+floor_count[7]['total_unit']))*100).toFixed(2)+"%"
                                                                                                                                               })
                                                                                                                                        })//35th    
                                                                                                                                    })//34th
                                                                                                                                })//33rd
                                                                                                                            })//32nd
                                                                                                                        })//31st
                                                                                                                    })//30th
                                                                                                                  })//29th
                                                                                                                })//28th
                                                                                                              })//27th
                                                                                                            })//26th
                                                                                                          })//25th
                                                                                                        })//24th
                                                                                                    })//23rd 
                                                                                                })//22nd
                                                                                            })//21st
                                                                                        })//20th
                                                                                    })//19th then
                                                                                })//18th then 
                                                                            })//17th then
                                                                        })//16th then
                                                                    })//15th then ending
                                                                })//14th then ending
                                                            })//13th then ending
                                                        })//12th then ending
                                                    })//11th then ending
                                                })//10th then ending
                                            })//9th then ending
                                        })//8th then ending
                                    })//7th then ending
                                });//6th then ending
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;        
}

//Excel function for Luxor 
function generateExcelLuxor(arr,arr15,arr16,arr17,arr18,arr19,arr20,arr21,arr22)
{
        //path to generate excel
        const path = require('path');
        let excel = new Excel(path.join(__dirname,'Luxor.xlsx'));
    
            // var tower = arr[0]['tower_name'];
           excel.writeSheet('StructureActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','TowerG','TowerH','CompletedPercent'],[arr]).then(()=>{
                     excel.writeRow('StructureActivities',1,{
                        Activity: "Unit Count",
                        TowerA : arr[13]['total_unit'],
                        TowerB : arr[14]['total_unit'],
                        TowerC : arr[15]['total_unit'],
                        TowerD : arr[16]['total_unit'],
                        TowerE : arr[17]['total_unit'],
                        TowerF : arr[18]['total_unit'],
                        TowerG : arr[19]['total_unit'],
                        TowerH : arr[20]['total_unit'],
                    }).then(()=>{
                            excel.writeRow('StructureActivities',2,{
                                Activity : arr15[0].milestone,
                                TowerA  : arr15[0].count,
                                TowerB  : arr16[0].count,
                                TowerC  : arr17[0].count,
                                TowerD  : arr18[0].count,
                                TowerE  : arr19[0].count,
                                TowerF :  arr20[0].count,
                                TowerG :  arr21[0].count,
                                TowerH :  arr22[0].count,        
                                CompletedPercent : ((arr15[0]['count']+arr16[0]['count']+arr17[0]['count']+arr18[0]['count']+arr19[0]['count']+arr20[0]['count']+arr21[0]['count']+arr22[0]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                            }).then(()=>{
                            excel.writeRow('StructureActivities',3,{
                                Activity : arr15[1].milestone,
                                TowerA  : arr15[1].count,
                                TowerB  : arr16[1].count,
                                TowerC  : arr17[1].count,
                                TowerD  : arr18[1].count,
                                TowerE  : arr19[1].count,
                                TowerF :  arr20[1].count,
                                TowerG :  arr21[1].count,
                                TowerH :  arr22[1].count,        
                                CompletedPercent : ((arr15[1]['count']+arr16[1]['count']+arr17[1]['count']+arr18[1]['count']+arr19[1]['count']+arr20[1]['count']+arr21[1]['count']+arr22[1]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                            }).then(()=>{
                                excel.writeRow('StructureActivities',4,{
                                    Activity : arr15[2].milestone,
                                    TowerA  : arr15[2].count,
                                    TowerB  : arr16[2].count,
                                    TowerC  : arr17[2].count,
                                    TowerD  : arr18[2].count,
                                    TowerE  : arr19[2].count,
                                    TowerF :  arr20[2].count,
                                    TowerG :  arr21[2].count,
                                    TowerH :  arr22[2].count,        
                                    CompletedPercent : ((arr15[2]['count']+arr16[2]['count']+arr17[2]['count']+arr18[2]['count']+arr19[2]['count']+arr20[2]['count']+arr21[2]['count']+arr22[2]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                    }).then(()=>{
                                    excel.writeRow('StructureActivities',5,{
                                        Activity : arr15[3].milestone,
                                        TowerA  : arr15[3].count,
                                        TowerB  : arr16[3].count,
                                        TowerC  : arr17[3].count,
                                        TowerD  : arr18[3].count,
                                        TowerE  : arr19[3].count,
                                        TowerF :  arr20[3].count,
                                        TowerG :  arr21[3].count,
                                        TowerH :  arr22[3].count,        
                                        CompletedPercent : ((arr15[3]['count']+arr16[3]['count']+arr17[3]['count']+arr18[3]['count']+arr19[3]['count']+arr20[3]['count']+arr21[3]['count']+arr22[3]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                        }).then(()=>{
                                        excel.writeRow('StructureActivities',6,{
                                            Activity : arr15[4].milestone,
                                            TowerA  : arr15[4].count,
                                            TowerB  : arr16[4].count,
                                            TowerC  : arr17[4].count,
                                            TowerD  : arr18[4].count,
                                            TowerE  : arr19[4].count,
                                            TowerF :  arr20[4].count,
                                            TowerG :  arr21[4].count,
                                            TowerH :  arr22[4].count,        
                                            CompletedPercent : ((arr15[4]['count']+arr16[4]['count']+arr17[4]['count']+arr18[4]['count']+arr19[4]['count']+arr20[4]['count']+arr21[4]['count']+arr22[4]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                            }).then(()=>{
                                            excel.writeRow('StructureActivities',7,{
                                                Activity : arr15[5].milestone,
                                                TowerA  : arr15[5].count,
                                                TowerB  : arr16[5].count,
                                                TowerC  : arr17[5].count,
                                                TowerD  : arr18[5].count,
                                                TowerE  : arr19[5].count,
                                                TowerF :  arr20[5].count,
                                                TowerG :  arr21[5].count,
                                                TowerH :  arr22[5].count,        
                                                CompletedPercent : ((arr15[5]['count']+arr16[5]['count']+arr17[5]['count']+arr18[5]['count']+arr19[5]['count']+arr20[5]['count']+arr21[5]['count']+arr22[5]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                                  }).then(()=>{
                                                excel.writeRow('StructureActivities',8,{
                                                    Activity : arr15[6].milestone,
                                                    TowerA  : arr15[6].count,
                                                    TowerB  : arr16[6].count,
                                                    TowerC  : arr17[6].count,
                                                    TowerD  : arr18[6].count,
                                                    TowerE  : arr19[6].count,
                                                    TowerF :  arr20[6].count,
                                                    TowerG :  arr21[6].count,
                                                    TowerH :  arr22[6].count,        
                                                    CompletedPercent : ((arr15[6]['count']+arr16[6]['count']+arr17[6]['count']+arr18[6]['count']+arr19[6]['count']+arr20[6]['count']+arr21[6]['count']+arr22[6]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                                      }).then(()=>{
                                                    excel.writeRow('StructureActivities',9,{
                                                        Activity : arr15[7].milestone,
                                                        TowerA  : arr15[7].count,
                                                        TowerB  : arr16[7].count,
                                                        TowerC  : arr17[7].count,
                                                        TowerD  : arr18[7].count,
                                                        TowerE  : arr19[7].count,
                                                        TowerF :  arr20[7].count,
                                                        TowerG :  arr21[7].count,
                                                        TowerH :  arr22[7].count,        
                                                        CompletedPercent : ((arr15[7]['count']+arr16[7]['count']+arr17[7]['count']+arr18[7]['count']+arr19[7]['count']+arr20[7]['count']+arr21[7]['count']+arr22[7]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"
                                                          }).then(()=>{
                                                        excel.writeRow('StructureActivities',10,{
                                                            Activity : arr15[8].milestone,
                                                            TowerA  : arr15[8].count,
                                                            TowerB  : arr16[8].count,
                                                            TowerC  : arr17[8].count,
                                                            TowerD  : arr18[8].count,
                                                            TowerE  : arr19[8].count,
                                                            TowerF :  arr20[8].count,
                                                            TowerG :  arr21[8].count,
                                                            TowerH :  arr22[8].count,        
                                                            CompletedPercent : ((arr15[8]['count']+arr16[8]['count']+arr17[8]['count']+arr18[8]['count']+arr19[8]['count']+arr20[8]['count']+arr21[8]['count']+arr22[8]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                      
                                                          }).then(()=>{
                                                            excel.writeRow('StructureActivities',11,{
                                                                Activity : arr15[9].milestone,
                                                                TowerA  : arr15[9].count,
                                                                TowerB  : arr16[9].count,
                                                                TowerC  : arr17[9].count,
                                                                TowerD  : arr18[9].count,
                                                                TowerE  : arr19[9].count,
                                                                TowerF :  arr20[9].count,
                                                                TowerG :  arr21[9].count,
                                                                TowerH :  arr22[9].count,        
                                                                CompletedPercent : ((arr15[9]['count']+arr16[9]['count']+arr17[9]['count']+arr18[9]['count']+arr19[9]['count']+arr20[9]['count']+arr21[9]['count']+arr22[9]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                      
                           
                                                              }).then(()=>{
                                                                excel.writeRow('StructureActivities',12,{
                                                                    Activity : arr15[10].milestone,
                                                                    TowerA  : arr15[10].count,
                                                                    TowerB  : arr16[10].count,
                                                                    TowerC  : arr17[10].count,
                                                                    TowerD  : arr18[10].count,
                                                                    TowerE  : arr19[10].count,
                                                                    TowerF :  arr20[10].count,
                                                                    TowerG :  arr21[10].count,
                                                                    TowerH :  arr22[10].count,        
                                                                    CompletedPercent : ((arr15[10]['count']+arr16[10]['count']+arr17[10]['count']+arr18[10]['count']+arr19[10]['count']+arr20[10]['count']+arr21[10]['count']+arr22[10]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                      
                                                                    }).then(()=>{
                                                                    excel.writeRow('StructureActivities',13,{
                                                                        Activity : arr15[11].milestone,
                                                                        TowerA  : arr15[11].count,
                                                                        TowerB  : arr16[11].count,
                                                                        TowerC  : arr17[11].count,
                                                                        TowerD  : arr18[11].count,
                                                                        TowerE  : arr19[11].count,
                                                                        TowerF :  arr20[11].count,
                                                                        TowerG :  arr21[11].count,
                                                                        TowerH :  arr22[11].count,        
                                                                        CompletedPercent : ((arr15[11]['count']+arr16[11]['count']+arr17[11]['count']+arr18[11]['count']+arr19[11]['count']+arr20[11]['count']+arr21[11]['count']+arr22[11]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                      
                                                                        }).then(()=>{
                                                                        excel.writeRow('StructureActivities',14,{
                                                                            Activity : arr15[12].milestone,
                                                                            TowerA  : arr15[12].count,
                                                                            TowerB  : arr16[12].count,
                                                                            TowerC  : arr17[12].count,
                                                                            TowerD  : arr18[12].count,
                                                                            TowerE  : arr19[12].count,
                                                                            TowerF :  arr20[12].count,
                                                                            TowerG :  arr21[12].count,
                                                                            TowerH :  arr22[12].count,        
                                                                            CompletedPercent : ((arr15[12]['count']+arr16[12]['count']+arr17[12]['count']+arr18[12]['count']+arr19[12]['count']+arr20[12]['count']+arr21[12]['count']+arr22[12]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                      
                                                                            }).then(()=>{
                                                                            excel.writeRow('StructureActivities',15,{
                                                                                Activity : arr15[13].milestone,
                                                                                TowerA  : arr15[13].count,
                                                                                TowerB  : arr16[13].count,
                                                                                TowerC  : arr17[13].count,
                                                                                TowerD  : arr18[13].count,
                                                                                TowerE  : arr19[13].count,
                                                                                TowerF :  arr20[13].count,
                                                                                TowerG :  arr21[13].count,
                                                                                TowerH :  arr22[13].count,        
                                                                                CompletedPercent : ((arr15[13]['count']+arr16[13]['count']+arr17[13]['count']+arr18[13]['count']+arr19[13]['count']+arr20[13]['count']+arr21[13]['count']+arr22[13]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                      
                                                                                }).then(()=>{
                                                                                excel.writeRow('StructureActivities',16,{
                                                                                    Activity : arr15[14].milestone,
                                                                                    TowerA  : arr15[14].count,
                                                                                    TowerB  : arr16[14].count,
                                                                                    TowerC  : arr17[14].count,
                                                                                    TowerD  : arr18[14].count,
                                                                                    TowerE  : arr19[14].count,
                                                                                    TowerF :  arr20[14].count,
                                                                                    TowerG :  arr21[14].count,
                                                                                    TowerH :  arr22[14].count,        
                                                                                    CompletedPercent : ((arr15[14]['count']+arr16[14]['count']+arr17[14]['count']+arr18[14]['count']+arr19[14]['count']+arr20[14]['count']+arr21[14]['count']+arr22[14]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                      
                                                                                                               
                                                                                }).then(()=>{
                                                                                    excel.writeRow('StructureActivities',17,{
                                                                                        Activity : arr15[15].milestone,
                                                                                        TowerA  : arr15[15].count,
                                                                                        TowerB  : arr16[15].count,
                                                                                        TowerC  : arr17[15].count,
                                                                                        TowerD  : arr18[15].count,
                                                                                        TowerE  : arr19[15].count,
                                                                                        TowerF :  arr20[15].count,
                                                                                        TowerG :  arr21[15].count,
                                                                                        TowerH :  arr22[15].count,        
                                                                                        CompletedPercent : ((arr15[15]['count']+arr16[15]['count']+arr17[15]['count']+arr18[15]['count']+arr19[15]['count']+arr20[15]['count']+arr21[15]['count']+arr22[15]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                                                                                                                                           
                                                                                    }).then(()=>{
                                                                                        excel.writeRow('StructureActivities',18,{
                                                                                            Activity : arr15[16].milestone,
                                                                                            TowerA  : arr15[16].count,
                                                                                            TowerB  : arr16[16].count,
                                                                                            TowerC  : arr17[16].count,
                                                                                            TowerD  : arr18[16].count,
                                                                                            TowerE  : arr19[16].count,
                                                                                            TowerF :  arr20[16].count,
                                                                                            TowerG :  arr21[16].count,
                                                                                            TowerH :  arr22[16].count,        
                                                                                            CompletedPercent : ((arr15[16]['count']+arr16[16]['count']+arr17[16]['count']+arr18[16]['count']+arr19[16]['count']+arr20[16]['count']+arr21[16]['count']+arr22[16]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                                                                                                                                           
                                                                                                                         
                                                                                        }).then(()=>{
                                                                                            excel.writeRow('StructureActivities',19,{
                                                                                                Activity : arr15[17].milestone,
                                                                                                TowerA  : arr15[17].count,
                                                                                                TowerB  : arr16[17].count,
                                                                                                TowerC  : arr17[17].count,
                                                                                                TowerD  : arr18[17].count,
                                                                                                TowerE  : arr19[17].count,
                                                                                                TowerF :  arr20[17].count,
                                                                                                TowerG :  arr21[17].count,
                                                                                                TowerH :  arr22[17].count,        
                                                                                                CompletedPercent : ((arr15[17]['count']+arr16[17]['count']+arr17[17]['count']+arr18[17]['count']+arr19[17]['count']+arr20[17]['count']+arr21[17]['count']+arr22[17]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                                                                                                                                           
                                                                                                                             
                                                                                            }).then(()=>{
                                                                                                excel.writeRow('StructureActivities',20,{
                                                                                                    Activity : arr15[18].milestone,
                                                                                                    TowerA  : arr15[18].count,
                                                                                                    TowerB  : arr16[18].count,
                                                                                                    TowerC  : arr17[18].count,
                                                                                                    TowerD  : arr18[18].count,
                                                                                                    TowerE  : arr19[18].count,
                                                                                                    TowerF :  arr20[18].count,
                                                                                                    TowerG :  arr21[18].count,
                                                                                                    TowerH :  arr22[18].count,        
                                                                                                    CompletedPercent : ((arr15[18]['count']+arr16[18]['count']+arr17[18]['count']+arr18[18]['count']+arr19[18]['count']+arr20[18]['count']+arr21[18]['count']+arr22[18]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                                                                                                                                           
                                                                                                                                
                                                                                                }).then(()=>{
                                                                                                    excel.writeRow('StructureActivities',21,{
                                                                                                        Activity : arr15[19].milestone,
                                                                                                        TowerA  : arr15[19].count,
                                                                                                        TowerB  : arr16[19].count,
                                                                                                        TowerC  : arr17[19].count,
                                                                                                        TowerD  : arr18[19].count,
                                                                                                        TowerE  : arr19[19].count,
                                                                                                        TowerF :  arr20[19].count,
                                                                                                        TowerG :  arr21[19].count,
                                                                                                        TowerH :  arr22[19].count,        
                                                                                                        CompletedPercent : ((arr15[19]['count']+arr16[19]['count']+arr17[19]['count']+arr18[19]['count']+arr19[19]['count']+arr20[19]['count']+arr21[19]['count']+arr22[19]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                                                                                                                                           
                                                                                                                                    
                                                                                                    }).then(()=>{
                                                                                                        excel.writeRow('StructureActivities',22,{
                                                                                                            Activity : arr15[20].milestone,
                                                                                                            TowerA  : arr15[20].count,
                                                                                                            TowerB  : arr16[20].count,
                                                                                                            TowerC  : arr17[20].count,
                                                                                                            TowerD  : arr18[20].count,
                                                                                                            TowerE  : arr19[20].count,
                                                                                                            TowerF :  arr20[20].count,
                                                                                                            TowerG :  arr21[20].count,
                                                                                                            TowerH :  arr22[20].count,        
                                                                                                            CompletedPercent : ((arr15[20]['count']+arr16[20]['count']+arr17[20]['count']+arr18[20]['count']+arr19[20]['count']+arr20[20]['count']+arr21[20]['count']+arr22[20]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                                                                                                                                           
                                                                                                                                       
                                                                                                        }).then(()=>{
                                                                                                            excel.writeRow('StructureActivities',23,{
                                                                                                                Activity : arr15[21].milestone,
                                                                                                                TowerA  : arr15[21].count,
                                                                                                                TowerB  : arr16[21].count,
                                                                                                                TowerC  : arr17[21].count,
                                                                                                                TowerD  : arr18[21].count,
                                                                                                                TowerE  : arr19[21].count,
                                                                                                                TowerF :  arr20[21].count,
                                                                                                                TowerG :  arr21[21].count,
                                                                                                                TowerH :  arr22[21].count,        
                                                                                                                CompletedPercent : ((arr15[21]['count']+arr16[21]['count']+arr17[21]['count']+arr18[21]['count']+arr19[21]['count']+arr20[21]['count']+arr21[21]['count']+arr22[21]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']+arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']))*100+"%"                                                                                                                                           
                                                                                                                    })                  
                                                                                                        });                                                                                                        
                                                                                                    });
                                                                                                });
                                                                                            });
                                                                                        });
                                                                                    });
                                                                                });
                                                                            });
                                                                        });  
                                                                    });
                                                                 });
                                                              });
                                                          });
                                                      });
                                                  });
                                              });
                                        });
                                    });
                                });
                            });
                        });
                    });
                });
            return;    
}


//Excel for greenField phse2
function generateExcelPhase2(garr,arr23,arr24,arr25)
{
    const path = require('path');
    let excel = new Excel(path.join(__dirname,'Greenfieldphase2.xlsx'));
    
    excel.writeSheet('StructureActivities',['Activity','TowerG','TowerH','TowerJ','CompletedPercent'],[garr]).then(()=>{
        excel.writeRow('StructureActivities',1,{
            Activity: "Unit Count",
            TowerG : garr[0]['total_unit'],
            TowerH : garr[1]['total_unit'],
            TowerJ : garr[2]['total_unit'],
        }).then(()=>{//1 then
                    excel.writeRow('StructureActivities',2,{
                        Activity : arr23[0].milestone,
                        TowerG : arr23[0]['count'],
                        TowerH : arr24[0]['count'],
                        TowerJ : arr25[0]['count'],
                        CompletedPercent : ((arr23[0]['count']+arr24[0]['count']+arr25[0]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                        }).then(()=>{//2nd then
                            excel.writeRow('StructureActivities',3,{
                                Activity : arr23[1].milestone,
                                TowerG : arr23[1]['count'],
                                TowerH : arr24[1]['count'],
                                TowerJ : arr25[1]['count'],
                                CompletedPercent : ((arr23[1]['count']+arr24[1]['count']+arr25[1]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                        }).then(()=>{//3rd then
                                    excel.writeRow('StructureActivities',4,{
                                        Activity : arr23[2].milestone,
                                        TowerG : arr23[2]['count'],
                                        TowerH : arr24[2]['count'],
                                        TowerJ : arr25[2]['count'],
                                        CompletedPercent : ((arr23[2]['count']+arr24[2]['count']+arr25[2]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                        }).then(()=>{//4th then
                                            excel.writeRow('StructureActivities',5,{
                                                Activity : arr23[3].milestone,
                                                TowerG : arr23[3]['count'],
                                                TowerH : arr24[3]['count'],
                                                TowerJ : arr25[3]['count'],
                                                CompletedPercent : ((arr23[3]['count']+arr24[3]['count']+arr25[3]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                  }).then(()=>{//5th then
                                                    excel.writeRow('StructureActivities',6,{
                                                        Activity : arr23[4].milestone,
                                                        TowerG : arr23[4]['count'],
                                                        TowerH : arr24[4]['count'],
                                                        TowerJ : arr25[4]['count'],
                                                        CompletedPercent : ((arr23[4]['count']+arr24[4]['count']+arr25[4]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                        }).then(()=>{//6th then
                                                            excel.writeRow('StructureActivities',7,{
                                                                Activity : arr23[5].milestone,
                                                                TowerG : arr23[5]['count'],
                                                                TowerH : arr24[5]['count'],
                                                                TowerJ : arr25[5]['count'],
                                                                CompletedPercent : ((arr23[5]['count']+arr24[5]['count']+arr25[5]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                    }).then(()=>{//7th then
                                                                    excel.writeRow('StructureActivities',8,{
                                                                        Activity : arr23[6].milestone,
                                                                        TowerG : arr23[6]['count'],
                                                                        TowerH : arr24[6]['count'],
                                                                        TowerJ : arr25[6]['count'],
                                                                        CompletedPercent : ((arr23[6]['count']+arr24[6]['count']+arr25[6]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                         }).then(()=>{//8th then
                                                                        excel.writeRow('StructureActivities',9,{
                                                                            Activity : arr23[7].milestone,
                                                                            TowerG : arr23[7]['count'],
                                                                            TowerH : arr24[7]['count'],
                                                                            TowerJ : arr25[7]['count'],
                                                                            CompletedPercent : ((arr23[7]['count']+arr24[7]['count']+arr25[7]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                           }).then(()=>{//9th then
                                                                            excel.writeRow('StructureActivities',10,{
                                                                                Activity : arr23[8].milestone,
                                                                                TowerG : arr23[8]['count'],
                                                                                TowerH : arr24[8]['count'],
                                                                                TowerJ : arr25[8]['count'],
                                                                                CompletedPercent : ((arr23[8]['count']+arr24[8]['count']+arr25[8]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                 }).then(()=>{//10th then
                                                                                excel.writeRow('StructureActivities',11,{
                                                                                    Activity : arr23[9].milestone,
                                                                                    TowerG : arr23[9]['count'],
                                                                                    TowerH : arr24[9]['count'],
                                                                                    TowerJ : arr25[9]['count'],
                                                                                    CompletedPercent : ((arr23[9]['count']+arr24[9]['count']+arr25[9]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                   }).then(()=>{//11th then
                                                                                    excel.writeRow('StructureActivities',12,{
                                                                                        Activity : arr23[10].milestone,
                                                                                        TowerG : arr23[10]['count'],
                                                                                        TowerH : arr24[10]['count'],
                                                                                        TowerJ : arr25[10]['count'],
                                                                                        CompletedPercent : ((arr23[10]['count']+arr24[10]['count']+arr25[10]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                        }).then(()=>{//12th then
                                                                                        excel.writeRow('StructureActivities',13,{
                                                                                            Activity : arr23[11].milestone,
                                                                                            TowerG : arr23[11]['count'],
                                                                                            TowerH : arr24[11]['count'],
                                                                                            TowerJ : arr25[11]['count'],
                                                                                            CompletedPercent : ((arr23[11]['count']+arr24[11]['count']+arr25[11]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                            }).then(()=>{//13th then
                                                                                            excel.writeRow('StructureActivities',14,{
                                                                                                Activity : arr23[12].milestone,
                                                                                                TowerG : arr23[12]['count'],
                                                                                                TowerH : arr24[12]['count'],
                                                                                                TowerJ : arr25[12]['count'],
                                                                                                CompletedPercent : ((arr23[12]['count']+arr24[12]['count']+arr25[12]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                             }).then(()=>{//14th then
                                                                                                excel.writeRow('StructureActivities',15,{
                                                                                                    Activity : arr23[13].milestone,
                                                                                                    TowerG : arr23[13]['count'],
                                                                                                    TowerH : arr24[13]['count'],
                                                                                                    TowerJ : arr25[13]['count'],
                                                                                                    CompletedPercent : ((arr23[13]['count']+arr24[13]['count']+arr25[13]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                        }).then(()=>{//15th then
                                                                                                    excel.writeRow('StructureActivities',16,{
                                                                                                        Activity : arr23[14].milestone,
                                                                                                        TowerG : arr23[14]['count'],
                                                                                                        TowerH : arr24[14]['count'],
                                                                                                        TowerJ : arr25[14]['count'],
                                                                                                        CompletedPercent : ((arr23[14]['count']+arr24[14]['count']+arr25[14]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                        }).then(()=>{//16th then
                                                                                                        excel.writeRow('StructureActivities',17,{
                                                                                                            Activity : arr23[15].milestone,
                                                                                                            TowerG : arr23[15]['count'],
                                                                                                            TowerH : arr24[15]['count'],
                                                                                                            TowerJ : arr25[15]['count'],
                                                                                                            CompletedPercent : ((arr23[15]['count']+arr24[15]['count']+arr25[15]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                            }).then(()=>{//17th then
                                                                                                            excel.writeRow('StructureActivities',18,{
                                                                                                                Activity : arr23[16].milestone,
                                                                                                                TowerG : arr23[16]['count'],
                                                                                                                TowerH : arr24[16]['count'],
                                                                                                                TowerJ : arr25[16]['count'],
                                                                                                                CompletedPercent : ((arr23[16]['count']+arr24[16]['count']+arr25[16]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                 }).then(()=>{//18th then
                                                                                                                excel.writeRow('StructureActivities',19,{
                                                                                                                    Activity : arr23[17].milestone,
                                                                                                                    TowerG : arr23[17]['count'],
                                                                                                                    TowerH : arr24[17]['count'],
                                                                                                                    TowerJ : arr25[17]['count'],
                                                                                                                    CompletedPercent : ((arr23[17]['count']+arr24[17]['count']+arr25[17]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                     }).then(()=>{ //19th then
                                                                                                                    excel.writeRow('StructureActivities',20,{
                                                                                                                        Activity : arr23[18].milestone,
                                                                                                                        TowerG : arr23[18]['count'],
                                                                                                                        TowerH : arr24[18]['count'],
                                                                                                                        TowerJ : arr25[18]['count'],
                                                                                                                        CompletedPercent : ((arr23[18]['count']+arr24[18]['count']+arr25[18]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                       }).then(()=>{//20th then
                                                                                                                        excel.writeRow('StructureActivities',21,{
                                                                                                                            Activity : arr23[19].milestone,
                                                                                                                            TowerG : arr23[19]['count'],
                                                                                                                            TowerH : arr24[19]['count'],
                                                                                                                            TowerJ : arr25[19]['count'],
                                                                                                                            CompletedPercent : ((arr23[19]['count']+arr24[19]['count']+arr25[19]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                            }).then(()=>{//21st then
                                                                                                                            excel.writeRow('StructureActivities',22,{
                                                                                                                                Activity : arr23[20].milestone,
                                                                                                                                TowerG : arr23[20]['count'],
                                                                                                                                TowerH : arr24[20]['count'],
                                                                                                                                TowerJ : arr25[20]['count'],
                                                                                                                                CompletedPercent : ((arr23[20]['count']+arr24[20]['count']+arr25[20]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                                }).then(()=>{//22nd then
                                                                                                                                excel.writeRow('StructureActivities',23,{
                                                                                                                                    Activity : arr23[21].milestone,
                                                                                                                                    TowerG : arr23[21]['count'],
                                                                                                                                    TowerH : arr24[21]['count'],
                                                                                                                                    TowerJ : arr25[21]['count'],
                                                                                                                                    CompletedPercent : ((arr23[21]['count']+arr24[21]['count']+arr25[21]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                                    }).then(()=>{//23rd then
                                                                                                                                    excel.writeRow('StructureActivities',24,{
                                                                                                                                        Activity : arr23[22].milestone,
                                                                                                                                        TowerG : arr23[22]['count'],
                                                                                                                                        TowerH : arr24[22]['count'],
                                                                                                                                        TowerJ : arr25[22]['count'],
                                                                                                                                        CompletedPercent : ((arr23[22]['count']+arr24[22]['count']+arr25[22]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                                            }).then(()=>{//24th then
                                                                                                                                        excel.writeRow('StructureActivities',25,{
                                                                                                                                            Activity : arr23[23].milestone,
                                                                                                                                            TowerG : arr23[23]['count'],
                                                                                                                                            TowerH : arr24[23]['count'],
                                                                                                                                            TowerJ : arr25[23]['count'],
                                                                                                                                            CompletedPercent : ((arr23[23]['count']+arr24[23]['count']+arr25[23]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                                            }).then(()=>{//25th then
                                                                                                                                            excel.writeRow('StructureActivities',26,{
                                                                                                                                                Activity : arr23[24].milestone,
                                                                                                                                                TowerG : arr23[24]['count'],
                                                                                                                                                TowerH : arr24[24]['count'],
                                                                                                                                                TowerJ : arr25[24]['count'],
                                                                                                                                                CompletedPercent : ((arr23[24]['count']+arr24[24]['count']+arr25[24]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                                               }).then(()=>{//26th then
                                                                                                                                                excel.writeRow('StructureActivities',27,{
                                                                                                                                                    Activity : arr23[25].milestone,
                                                                                                                                                    TowerG : arr23[25]['count'],
                                                                                                                                                    TowerH : arr24[25]['count'],
                                                                                                                                                    TowerJ : arr25[25]['count'],
                                                                                                                                                    CompletedPercent : ((arr23[25]['count']+arr24[25]['count']+arr25[25]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                                                            }).then(()=>{//27th then
                                                                                                                                                    excel.writeRow('StructureActivities',28,{
                                                                                                                                                        Activity : arr23[26].milestone,
                                                                                                                                                        TowerG : arr23[26]['count'],
                                                                                                                                                        TowerH : arr24[26]['count'],
                                                                                                                                                        TowerJ : arr25[26]['count'],
                                                                                                                                                        CompletedPercent : ((arr23[26]['count']+arr24[26]['count']+arr25[26]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                                                         }).then(()=>{//28th then
                                                                                                                                                            excel.writeRow('StructureActivities',29,{
                                                                                                                                                                Activity : arr23[27].milestone,
                                                                                                                                                                TowerG : arr23[27]['count'],
                                                                                                                                                                TowerH : arr24[27]['count'],
                                                                                                                                                                TowerJ : arr25[27]['count'],
                                                                                                                                                                CompletedPercent : ((arr23[27]['count']+arr24[27]['count']+arr25[27]['count'])/(garr[0]['total_unit']+garr[1]['total_unit']+garr[2]['total_unit']))*100+"%",
                                                                                                                                                                 })

                                                                                                                                                         });//28th then ending
                                                                                                                                                });//27th then ending
                                                                                                                                            });//26th then ending
                                                                                                                                        });//25th then ending
                                                                                                                                    })//24th then ending
                                                                                                                                })//23rd then ending  
                                                                                                                            });//22nd then ending
                                                                                                                        });//21st then ending
                                                                                                                    });//20th then ending
                                                                                                                });//19th then ending
                                                                                                            });//18th then ending
                                                                                                        });//17th then ending
                                                                                                    })//16th then ending
                                                                                                });//15th then ending
                                                                                            });//14th then ending
                                                                                        });//13th then ending
                                                                                    });//12th then ending
                                                                                });//11th then ending
                                                                            });//10th then ending
                                                                        });//9th then ending
                                                                    });//8th then ending
                                                                });//7th then ending
                                                        });//6th then ending
                                                })//5th then ending
                                        });//4th then ending
                                });//3rd then ending  
                        });//2nd then ending
                    });//1st then ending
    });//excel writesheet
    return;
}

//other Activiyty excel for panorama hills
function generateExcelOtherPanorama(arr,Tower2A,Tower2B,Tower2C,Tower2D)
{

            //path to generate excel
            const path = require('path');
            let excel = new Excel(path.join(__dirname,'PanoramaHills.xlsx'));
    
            excel.writeSheet('OtherActivities',['Activity','Tower2A','Tower2B','Tower2C','Tower2D','CompletedPercent'],[arr]).then(()=>{
                excel.writeRow('OtherActivities',1,{
                   Activity: "Floor Count",
                   Tower2A : arr[0]['total_unit'],
                   Tower2B : arr[1]['total_unit'],
                   Tower2C : arr[2]['total_unit'],
                   Tower2D : arr[3]['total_unit'],
               }).then(()=>{//1st then
                   excel.writeRow('OtherActivities',2,{
                    Activity: Tower2A[0].milestone,
                    Tower2A : Tower2A[0]['count'],
                    Tower2B : Tower2B[0]['count'],
                    Tower2C : Tower2C[0]['count'],
                    Tower2D : Tower2C[0]['count'],
                    CompletedPercent : ((Tower2A[0]['count']+Tower2B[0]['count']+Tower2C[0]['count']+Tower2D[0]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"       
                 }).then(()=>{//2nd then
                    excel.writeRow('OtherActivities',3,{
                        Activity: Tower2A[1].milestone,
                        Tower2A : Tower2A[1]['count'],
                        Tower2B : Tower2B[1]['count'],
                        Tower2C : Tower2C[1]['count'],
                        Tower2D : Tower2D[1]['count'],
                        CompletedPercent : ((Tower2A[1]['count']+Tower2B[1]['count']+Tower2C[1]['count']+Tower2D[1]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                          
                    }).then(()=>{//3rd then
                        excel.writeRow('OtherActivities',4,{
                            Activity: Tower2A[2].milestone,
                            Tower2A : Tower2A[2]['count'],
                            Tower2B : Tower2B[2]['count'],
                            Tower2C : Tower2C[2]['count'],
                            Tower2D : Tower2D[2]['count'],
                            CompletedPercent : ((Tower2A[2]['count']+Tower2B[2]['count']+Tower2C[2]['count']+Tower2D[2]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                   
                        }).then(()=>{//4th then
                            excel.writeRow('OtherActivities',5,{
                                Activity: Tower2A[3].milestone,
                                Tower2A : Tower2A[3]['count'],
                                Tower2B : Tower2B[3]['count'],
                                Tower2C : Tower2C[3]['count'],
                                Tower2D : Tower2D[3]['count'],
                                CompletedPercent : ((Tower2A[3]['count']+Tower2B[3]['count']+Tower2C[3]['count']+Tower2D[3]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                
                            }).then(()=>{//5th then
                                excel.writeRow('OtherActivities',6,{
                                    Activity: Tower2A[4].milestone,
                                    Tower2A : Tower2A[4]['count'],
                                    Tower2B : Tower2B[4]['count'],
                                    Tower2C : Tower2C[4]['count'],
                                    Tower2D : Tower2D[4]['count'],
                                    CompletedPercent : ((Tower2A[4]['count']+Tower2B[4]['count']+Tower2C[4]['count']+Tower2D[4]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                 
                                }).then(()=>{//6th then
                                    excel.writeRow('OtherActivities',7,{
                                        Activity: Tower2A[5].milestone,
                                        Tower2A : Tower2A[5]['count'],
                                        Tower2B : Tower2B[5]['count'],
                                        Tower2C : Tower2C[5]['count'],
                                        Tower2D : Tower2C[5]['count'],                     
                                        CompletedPercent : ((Tower2A[5]['count']+Tower2B[5]['count']+Tower2C[5]['count']+Tower2D[5]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                 
                                    }).then(()=>{//7th then
                                        excel.writeRow('OtherActivities',8,{
                                            Activity: Tower2A[6].milestone,
                                            Tower2A : Tower2A[6]['count'],
                                            Tower2B : Tower2B[6]['count'],
                                            Tower2C : Tower2C[6]['count'],
                                            Tower2D : Tower2C[6]['count'],                     
                                            CompletedPercent : ((Tower2A[6]['count']+Tower2B[6]['count']+Tower2C[6]['count']+Tower2D[6]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                 

                                        }).then(()=>{
                                            excel.writeRow('OtherActivities',9,{
                                                Activity: Tower2A[7].milestone,
                                                Tower2A : Tower2A[7]['count'],
                                                Tower2B : Tower2B[7]['count'],
                                                Tower2C : Tower2C[7]['count'],
                                                Tower2D : Tower2C[7]['count'],                     
                                                CompletedPercent : ((Tower2A[7]['count']+Tower2B[7]['count']+Tower2C[7]['count']+Tower2D[7]['count'])/(arr[0]['total_unit']+arr[1]['total_unit']+arr[2]['total_unit']+arr[3]['total_unit']))*100+"%"                                                                 
                                            });
                                            });//8th then endning
                                        });//7th then ending
                                });//6th then ending
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;        
}

//generate excel for phase 1
function generateExcelOtherPhase1(arr,arr6,arr7,arr8,arr9,arr10,arr11)
{
            //path to generate excel
            const path = require('path');
            let excel = new Excel(path.join(__dirname,'Greenfieldphase1.xlsx'));
    
            excel.writeSheet('OtherActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','CompletedPercent'],[arr]).then(()=>{
                excel.writeRow('OtherActivities',1,{
                   Activity: "Floor Count",
                   TowerA : arr[4]['total_unit'],
                   TowerB : arr[5]['total_unit'],
                   TowerC : arr[6]['total_unit'],
                   TowerD : arr[7]['total_unit'],
                   TowerE : arr[8]['total_unit'],
                   TowerF : arr[9]['total_unit']
               }).then(()=>{//1st then
                   excel.writeRow('OtherActivities',2,{
                    Activity: arr6[0].milestone,
                    TowerA : arr6[0]['count'],
                    TowerB : arr7[0]['count'],
                    TowerC : arr8[0]['count'],
                    TowerD : arr9[0]['count'],
                    TowerE : arr10[0]['count'],
                    TowerF : arr11[0]['count'],
                    CompletedPercent : (((arr6[0]['count']+arr7[0]['count']+arr8[0]['count']+arr9[0]['count']+arr10[0]['count']+arr11[0]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100)+"%"
                 }).then(()=>{//2nd then
                    excel.writeRow('OtherActivities',3,{
                        Activity: arr6[1].milestone,
                        TowerA : arr6[1]['count'],
                        TowerB : arr7[1]['count'],
                        TowerC : arr8[1]['count'],
                        TowerD : arr9[1]['count'],
                        TowerE : arr10[1]['count'],
                        TowerF : arr11[1]['count'],
                        CompletedPercent : ((arr6[1]['count']+arr7[1]['count']+arr8[1]['count']+arr9[1]['count']+arr10[1]['count']+arr11[1]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%"       
                        }).then(()=>{//3rd then
                        excel.writeRow('OtherActivities',4,{
                            Activity: arr6[2].milestone,
                            TowerA : arr6[2]['count'],
                            TowerB : arr7[2]['count'],
                            TowerC : arr8[2]['count'],
                            TowerD : arr9[2]['count'],
                            TowerE : arr10[2]['count'],
                            TowerF : arr11[2]['count'],
                            CompletedPercent : ((arr6[2]['count']+arr7[2]['count']+arr8[2]['count']+arr9[2]['count']+arr10[2]['count']+arr11[2]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%"       
                            }).then(()=>{//4th then
                            excel.writeRow('OtherActivities',5,{
                                Activity: arr6[3].milestone,
                                TowerA : arr6[3]['count'],
                                TowerB : arr7[3]['count'],
                                TowerC : arr8[3]['count'],
                                TowerD : arr9[3]['count'],
                                TowerE : arr10[3]['count'],
                                TowerF : arr11[3]['count'],
                                CompletedPercent : ((arr6[3]['count']+arr7[3]['count']+arr8[3]['count']+arr9[3]['count']+arr10[3]['count']+arr11[3]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%"       
                             }).then(()=>{//5th then
                                excel.writeRow('OtherActivities',6,{
                                    Activity: arr6[4].milestone,
                                    TowerA : arr6[4]['count'],
                                    TowerB : arr7[4]['count'],
                                    TowerC : arr8[4]['count'],
                                    TowerD : arr9[4]['count'],
                                    TowerE : arr10[4]['count'],
                                    TowerF : arr11[4]['count'],
                                    CompletedPercent : ((arr6[4]['count']+arr7[4]['count']+arr8[4]['count']+arr9[4]['count']+arr10[4]['count']+arr11[4]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%"       
                                    }).then(()=>{//6th then
                                    excel.writeRow('OtherActivities',7,{
                                        Activity: arr6[5].milestone,
                                        TowerA : arr6[5]['count'],
                                        TowerB : arr7[5]['count'],
                                        TowerC : arr8[5]['count'],
                                        TowerD : arr9[5]['count'],
                                        TowerE : arr10[5]['count'],
                                        TowerF : arr11[5]['count'],
                                        CompletedPercent : ((arr6[5]['count']+arr7[5]['count']+arr8[5]['count']+arr9[5]['count']+arr10[5]['count']+arr11[5]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%"       
                                        }).then(()=>{//7th then
                                        excel.writeRow('OtherActivities',8,{
                                            Activity: "External Painting",
                                            TowerA : arr6[6]['count'],
                                            TowerB : arr7[6]['count'],
                                            TowerC : arr8[6]['count'],
                                            TowerD : arr9[6]['count'],
                                            TowerE : arr10[6]['count'],
                                            TowerF : arr11[6]['count'],
                                            CompletedPercent : ((arr6[6]['count']+arr7[6]['count']+arr8[6]['count']+arr9[6]['count']+arr10[6]['count']+arr11[6]['count'])/(arr[4]['total_unit']+arr[5]['total_unit']+arr[6]['total_unit']+arr[7]['total_unit']+arr[8]['total_unit']+arr[9]['total_unit']))*100+"%",              
                                        })
                                        });//7th then ending
                                });//6th then ending
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;        
}



//generate excel for phase 2
function generateExcelOtherPhase2(arr,arr23,arr24,arr25)
{
            //path to generate excel
            const path = require('path');
            let excel = new Excel(path.join(__dirname,'GreenFieldPhase2.xlsx'));
    
            excel.writeSheet('OtherActivities',['Activity','TowerG','TowerH','TowerJ','CompletedPercent'],[arr]).then(()=>{
                excel.writeRow('OtherActivities',1,{
                   Activity: "Floor Count",
                   TowerG : arr[10]['total_unit'],
                   TowerH : arr[11]['total_unit'],
                   TowerJ : arr[12]['total_unit'],
               }).then(()=>{//1st then
                   excel.writeRow('OtherActivities',2,{
                    Activity: arr23[0].milestone,
                    TowerG : arr23[0]['count'],
                    TowerH : arr24[0]['count'],
                    TowerJ : arr25[0]['count'],
                    CompletedPercent : ((arr23[0]['count']+arr24[0]['count']+arr25[0]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"
                 }).then(()=>{//2nd then
                    excel.writeRow('OtherActivities',3,{
                        Activity: arr23[1].milestone,
                        TowerG : arr23[1]['count'],
                        TowerH : arr24[1]['count'],
                        TowerJ : arr25[1]['count'],
                        CompletedPercent : ((arr23[1]['count']+arr24[1]['count']+arr25[1]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"
                        }).then(()=>{//3rd then
                        excel.writeRow('OtherActivities',4,{
                            Activity: arr23[2].milestone,
                            TowerG : arr23[2]['count'],
                            TowerH : arr24[2]['count'],
                            TowerJ : arr25[2]['count'],
                            CompletedPercent : ((arr23[2]['count']+arr24[2]['count']+arr25[2]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"
                            }).then(()=>{//4th then
                            excel.writeRow('OtherActivities',5,{
                                Activity: arr23[3].milestone,
                                TowerG : arr23[3]['count'],
                                TowerH : arr24[3]['count'],
                                TowerJ : arr25[3]['count'],
                                CompletedPercent : ((arr23[3]['count']+arr24[3]['count']+arr25[3]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"
                                 }).then(()=>{//5th then
                                excel.writeRow('OtherActivities',6,{
                                    Activity: "External Painting",
                                    TowerG : arr23[4]['count'],
                                    TowerH : arr24[4]['count'],
                                    TowerJ : arr25[4]['count'],
                                    CompletedPercent : ((arr23[4]['count']+arr24[4]['count']+arr25[4]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"
                                   })
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;        
}


//function to generate excel for southren crest
function generateExcelOtherSC(arr,arr12,arr13,arr14)
{
            //path to generate excel
            const path = require('path');
            let excel = new Excel(path.join(__dirname,'SouthrenCrest.xlsx'));
    
            excel.writeSheet('OtherActivities',['Activity','TowerA','TowerB','TowerC','CompletedPercent'],[arr]).then(()=>{
                excel.writeRow('OtherActivities',1,{
                   Activity: "Floor Count",
                   TowerA : arr[13]['total_unit'],
                   TowerB : arr[14]['total_unit'],
                   TowerC : arr[15]['total_unit'],
               }).then(()=>{//1st then
                   excel.writeRow('OtherActivities',2,{
                    Activity: arr12[0].milestone,
                    TowerA : arr12[0]['count'],
                    TowerB : arr13[0]['count'],
                    TowerC : arr14[0]['count'],
                    CompletedPercent : ((arr12[0]['count']+arr13[0]['count']+arr14[0]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']))*100+"%"
                 }).then(()=>{//2nd then
                    excel.writeRow('OtherActivities',3,{
                        Activity: arr12[1].milestone,
                        TowerA : arr12[1]['count'],
                        TowerB : arr13[1]['count'],
                        TowerC : arr14[1]['count'],
                        CompletedPercent : ((arr12[1]['count']+arr13[1]['count']+arr14[1]['count'])/(arr[13]['total_unit']+arr[14]['total_unit']+arr[15]['total_unit']))*100+"%"
                        }).then(()=>{//3rd then
                        excel.writeRow('OtherActivities',4,{
                            Activity: arr12[2].milestone,
                            TowerA : arr12[2]['count'],
                            TowerB : arr13[2]['count'],
                            TowerC : arr14[2]['count'],
                            CompletedPercent : ((arr12[2]['count']+arr13[2]['count']+arr14[2]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"
                            }).then(()=>{//4th then
                            excel.writeRow('OtherActivities',5,{
                                Activity: arr12[3].milestone,
                                TowerA : arr12[3]['count'],
                                TowerB : arr13[3]['count'],
                                TowerC : arr14[3]['count'],
                                CompletedPercent : ((arr12[3]['count']+arr13[3]['count']+arr14[3]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"
                                     }).then(()=>{//5th then
                                excel.writeRow('OtherActivities',6,{
                                    Activity: arr12[4].milestone,
                                    TowerA : arr12[4]['count'],
                                    TowerB : arr13[4]['count'],
                                    TowerC : arr14[4]['count'],
                                    CompletedPercent : ((arr12[4]['count']+arr13[4]['count']+arr14[4]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"
                                    }).then(() => {//6th then
                                        excel.writeRow('OtherActivities',7,{
                                        Activity: arr12[5].milestone,
                                        TowerA : arr12[5]['count'],
                                        TowerB : arr13[5]['count'],
                                        TowerC : arr14[5]['count'],
                                        CompletedPercent : ((arr12[5]['count']+arr13[5]['count']+arr14[5]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"    
                                     }).then(() => {//7th then
                                        excel.writeRow('OtherActivities',8,{
                                            Activity: "External Painting",
                                            TowerA : arr12[6]['count'],
                                            TowerB : arr13[6]['count'],
                                            TowerC : arr14[6]['count'],
                                            CompletedPercent : ((arr12[6]['count']+arr13[6]['count']+arr14[6]['count'])/(arr[10]['total_unit']+arr[11]['total_unit']+arr[12]['total_unit']))*100+"%"        
                                      })
                                    });//7th then ending
                                });//6th then ending
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;        
}


//fucntion to generate excel for luxor
function generateExcelOtherLuxor(arr,arr15,arr16,arr17,arr18,arr19,arr20,arr21,arr22)
{
            //path to generate excel
            const path = require('path');
            let excel = new Excel(path.join(__dirname,'Luxor.xlsx'));
    
            excel.writeSheet('OtherActivities',['Activity','TowerA','TowerB','TowerC','TowerD','TowerE','TowerF','TowerG','TowerH','CompletedPercent'],[arr]).then(()=>{
                excel.writeRow('OtherActivities',1,{
                   Activity: "Floor Count",
                   TowerA : arr[16]['total_unit'],
                   TowerB : arr[17]['total_unit'],
                   TowerC : arr[18]['total_unit'],
                   TowerD : arr[19]['total_unit'],
                   TowerE : arr[20]['total_unit'],
                   TowerF : arr[21]['total_unit'],
                   TowerG : arr[22]['total_unit'],
                   TowerH : arr[23]['total_unit'],
               }).then(()=>{//1st then
                   excel.writeRow('OtherActivities',2,{
                    Activity: arr15[0].milestone,
                    TowerA : arr15[0]['count'],
                    TowerB : arr16[0]['count'],
                    TowerC : arr17[0]['count'],
                    TowerD : arr18[0]['count'],
                    TowerE : arr19[0]['count'],
                    TowerF : arr20[0]['count'],
                    TowerG : arr21[0]['count'],
                    TowerH : arr22[0]['count'],
                    CompletedPercent : ((arr15[0]['count']+arr16[0]['count']+arr17[0]['count']+arr18[0]['count']+arr19[0]['count']+arr20[0]['count']+arr21[0]['count']+arr22[0]['count'])/(arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']+arr[21]['total_unit']+arr[22]['total_unit']+arr[23]['total_unit']))*100+"%"
                 }).then(()=>{//2nd then
                    excel.writeRow('OtherActivities',3,{
                        Activity: arr15[1].milestone,
                        TowerA : arr15[1]['count'],
                        TowerB : arr16[1]['count'],
                        TowerC : arr17[1]['count'],
                        TowerD : arr18[1]['count'],
                        TowerE : arr19[1]['count'],
                        TowerF : arr20[1]['count'],
                        TowerG : arr21[1]['count'],
                        TowerH : arr22[1]['count'],
                        CompletedPercent : ((arr15[1]['count']+arr16[1]['count']+arr17[1]['count']+arr18[1]['count']+arr19[1]['count']+arr20[1]['count']+arr21[1]['count']+arr22[1]['count'])/(arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']+arr[21]['total_unit']+arr[22]['total_unit']+arr[23]['total_unit']))*100+"%"
                        }).then(()=>{//3rd then
                        excel.writeRow('OtherActivities',4,{
                            Activity: arr15[2].milestone,
                            TowerA : arr15[2]['count'],
                            TowerB : arr16[2]['count'],
                            TowerC : arr17[2]['count'],
                            TowerD : arr18[2]['count'],
                            TowerE : arr19[2]['count'],
                            TowerF : arr20[2]['count'],
                            TowerG : arr21[2]['count'],
                            TowerH : arr22[2]['count'],
                            CompletedPercent : ((arr15[2]['count']+arr16[2]['count']+arr17[2]['count']+arr18[2]['count']+arr19[2]['count']+arr20[2]['count']+arr21[2]['count']+arr22[2]['count'])/(arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']+arr[21]['total_unit']+arr[22]['total_unit']+arr[23]['total_unit']))*100+"%"
                            }).then(()=>{//4th then
                            excel.writeRow('OtherActivities',5,{
                                Activity: arr15[3].milestone,
                                TowerA : arr15[3]['count'],
                                TowerB : arr16[3]['count'],
                                TowerC : arr17[3]['count'],
                                TowerD : arr18[3]['count'],
                                TowerE : arr19[3]['count'],
                                TowerF : arr20[3]['count'],
                                TowerG : arr21[3]['count'],
                                TowerH : arr22[3]['count'],
                                CompletedPercent : ((arr15[3]['count']+arr16[3]['count']+arr17[3]['count']+arr18[3]['count']+arr19[3]['count']+arr20[3]['count']+arr21[3]['count']+arr22[3]['count'])/(arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']+arr[21]['total_unit']+arr[22]['total_unit']+arr[23]['total_unit']))*100+"%"
                                }).then(()=>{//5th then
                                excel.writeRow('OtherActivities',6,{
                                    Activity: arr15[4].milestone,
                                    TowerA : arr15[4]['count'],
                                    TowerB : arr16[4]['count'],
                                    TowerC : arr17[4]['count'],
                                    TowerD : arr18[4]['count'],
                                    TowerE : arr19[4]['count'],
                                    TowerF : arr20[4]['count'],
                                    TowerG : arr21[4]['count'],
                                    TowerH : arr22[4]['count'],
                                    CompletedPercent : ((arr15[4]['count']+arr16[4]['count']+arr17[4]['count']+arr18[4]['count']+arr19[4]['count']+arr20[4]['count']+arr21[4]['count']+arr22[4]['count'])/(arr[16]['total_unit']+arr[17]['total_unit']+arr[18]['total_unit']+arr[19]['total_unit']+arr[20]['total_unit']+arr[21]['total_unit']+arr[22]['total_unit']+arr[23]['total_unit']))*100+"%"
                                    })
                            });//5th then ening
                        });//4th then ending                   
                    });//3rd then ending   
                 });//2nd then ending
               });//1st then ending
            });
      return;        
}

module.exports = router;