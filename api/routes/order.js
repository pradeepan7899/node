const express = require('express');
const router  = express.Router();

router.get('/',(req, res, next)=>{
    res.status(200).json({
           message: 'get request' 
    });
});

router.get('/:id',(req, res)=>{
    var id = req.params.id;
    console.log(id);

    res.status(200).json({
        error   : false,
        message : 'Order get request',
        Orderid : id
    });
});


module.exports = router;