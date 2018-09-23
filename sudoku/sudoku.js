var arr = [[1,2,3,4,5,6,7,8,9],[1,2,3,4,5,6,7,8,10],[1,2,3,4,5,6,7,8,9],[1,2,3,4,5,6,7,8,9],[1,2,3,4,5,6,7,8,9],[1,2,3,4,5,6,7,8,9],[1,2,3,4,5,6,7,8,9],[1,2,3,4,5,6,7,8,9],[1,2,3,4,5,6,7,8,9]];
var checkrow = checkRow(arr);
var checkcolumn = checkColumn(arr);
var checkgrid = checkGrid(arr);

if(checkrow == true && checkcolumn == true && checkgrid == true)
{
    console.log("sudoku completed");
}
else
{
    console.log("sudoku is incomplete");    
}

/*
 *To Check Duplicate elements in Row
 */
 function checkRow(arr) {

     var rowArr = new Array();

     for(var i=0;i<arr.length;i++)//it traverses columns one after another
     {
        var rowSum = 0;

        for(var j=0;j<arr.length;j++)//it traverses thru row of each column
        {
            rowSum += arr[i][j]; 
        }

        //to check weather sum of each
        // row is equal to 45  
        rowSum == 45 ? rowArr.push(true) : rowArr.push(false); 
     }
     return (rowArr.indexOf(false) == -1 ? true : false);         
 }


/*
 *To Check Duplicate elements in column 
 */
 function checkColumn(arr) { 
    var colArr = new Array();

    for(var i=0;i<arr.length;i++)//it traverses columns one after another
    {
       var colSum = 0;

       for(var j=0;j<arr.length;j++)//it traverses thru row of each column
       {
            colSum += arr[j][i]; 
       }

       //to check weather sum of each
       // row is equal to 45  
       colSum == 45 ? colArr.push(true) : colArr.push(false); 
    }
    return (colArr.indexOf(false) == -1 ? true : false);         
 }

 /*
  *To Check Duplicate elements in 3x3 grid's
  */
  function checkGrid(arr) {
    var gridArr = new Array();    

        //1st grid
        var sum = 0;
        for(i=0;i<arr.length/3;i++)
        {
            for(j=0;j<arr.length/3;j++)
            {
                sum += arr[i][j];
            }
        }
        sum == 45 ? gridArr.push(true) : gridArr.push(false);

        //2nd grid
        var sum = 0;
        for(i=0;i<arr.length/3;i++)
        {
            for(j=3;j<6;j++)
            {
                sum += arr[i][j];
            }
        }
        sum == 45 ? gridArr.push(true) : gridArr.push(false);

        //3rd grid
        var sum = 0;
        for(i=0;i<arr.length/3;i++)
        {
          for(j=6;j<9;j++)
          {
               sum += arr[i][j];
          }
        }
        sum == 45 ? gridArr.push(true) : gridArr.push(false);
        
        //4th grid
        var sum = 0;
        for(i=3;i<6;i++)
        {
            for(j=0;j<arr.length/3;j++)
            {
                sum += arr[i][j];
            }
        }
        sum == 45 ? gridArr.push(true) : gridArr.push(false);

        
        //5th grid
        var sum = 0;
        for(i=3;i<6;i++)
        {
            for(j=3;j<6;j++)
            {
                sum += arr[i][j];
            }
        }
        sum == 45 ? gridArr.push(true) : gridArr.push(false);

        //6th grid
        var sum = 0;
        for(i=3;i<6;i++)
        {
            for(j=6;j<9;j++)
            {
                sum += arr[i][j];
            }
        }
        sum == 45 ? gridArr.push(true) : gridArr.push(false);
        
        //7th grid
        var sum = 0;
        for(i=6;i<9;i++)
        {
            for(j=0;j<3;j++)
            {
                sum += arr[i][j];
            }
        }
        sum == 45 ? gridArr.push(true) : gridArr.push(false);

        //8th grid
        var sum = 0;
        for(i=6;i<9;i++)
        {
            for(j=3;j<6;j++)
            {
                sum += arr[i][j];
            }
        }
        sum == 45 ? gridArr.push(true) : gridArr.push(false);

        //9th grid
        var sum = 0;
        for(i=6;i<9;i++)
        {
            for(j=6;j<9;j++)
            {
                sum += arr[i][j];
            }
        }
        sum == 45 ? gridArr.push(true) : gridArr.push(false);
        
    return (gridArr.indexOf(false) == -1 ? true : false);     
}
