const express = require('express');
const app     = express();
const morgan  = require('morgan');
const bodyParser = require('body-parser');


const loginRoute    = require('./api/routes/authenticate'); 

//loging
app.use(morgan('dev'));

//bodyParser
app.use(bodyParser.urlencoded({extended: false}));
app.use(bodyParser.json());


//handling COARS Error
app.use((req, res, next)=>{
    res.header('Access-Control-Allow-Origin','*');
    res.header('Access-control-allow-Headers','Origin, X-Requested-With, Content-Type, Accept, Authorization')
   
    if(req.method === 'OPTIONS'){
        res.header('Access-Control-Allow-Methods','PUT, POST, DELETE, GET');
        return res.status(200).json({});
    }
    next();
});


//routes to handle requests
app.use('/userauthenticate',loginRoute)

//error handling middleware
app.use((req, res, next)=>{
    const error = new Error('invalid Endpoint');
    error.status = 404;
    next(error);
});

app.use((error, req, res, next )=> {
    res.status(error.status || 500);
    res.json({
       error  : true, 
       message: error.message
    });
});


module.exports = app;