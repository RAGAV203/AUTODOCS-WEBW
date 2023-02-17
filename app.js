//jshint eversion:6

const fs = require('fs');
const express = require("express");
const bodyparser = require("body-parser");
const upload = require('express-fileupload')
const app = express();
const Docx = require('docx');
const path = require('path');
const mammoth = require('mammoth');
app.use(bodyparser.urlencoded({extended: true}));
app.use(express.static("public"));
app.use(upload())
app.set('view engine','ejs');
app.get("/",function(req,res){
    res.render("login");
});
app.get("/admin",function(req,res){
    res.render("admin");
});
const router = express.Router();

let myArray = [];
var filename1="";
app.post("/admin",function(req,res){
  if(req.files){
        var typeName = req.body.typename;
        console.log('Type Name: ', typeName);
        console.log(req.files)
        myArray.push(req.body.typename);
        var file = req.files.file
        filename1=file.name
        var typename=req.files.typename
        console.log(filename1)
        const filePath = './uploads/' + filename1;
        const fileData = req.files.file.data;
        console.log(myArray);
  // Use mammoth to extract the text from the Word document
  let matches = [];
  mammoth.extractRawText({ buffer: fileData })
    .then(result => {
        const regex = /\((.*?)\)/g;
        let match;
        while (match = regex.exec(result.value)) {
              matches.push(match[1]);
      }
      // Convert the labels array to an object with empty values
  const data = {};
  for (let i = 0; i < matches.length; i++) {
   data[matches[i]] = "";
}

const filePath2 = __dirname+'/uploads/'+filename1+'.json';

fs.writeFile(filePath2, JSON.stringify(data), (err) => {
 if (err) {
   console.error(err);
   res.status(500).send('Error writing to file');
   return;
 }
 console.log('The file has been saved!');
});

   file.mv('./uploads/'+filename1,function(err)
   {
       if(err)
       {
           res.send(err)
       }
       else{
           const nlabels = JSON.parse(fs.readFileSync('./uploads/'+filename1+'.json'),'utf8')
           res.render('adminform', { nlabels });
       }
   })
    }).catch(error => {
      console.error(error);
      res.status(500).send('Error extracting text from file');
    })
    }
});


app.get("/adminform",function(req,res){
    res.render('adminform')
});

app.get("/user",function(req,res){
  res.render('user', { myArray });
});
let selectedValue ="";
app.post('/user', (req, res) => {
   selectedValue = req.body.dropdown;
  // Store the selected value in a variable or database
  console.log('Selected value:', selectedValue);
  res.redirect('/userform');
});
app.post("/adminform",function(req,res){
  const formData = req.body;
  
  // Write the form data to a JSON file
  fs.writeFile('./uploads/'+filename1+'.json', JSON.stringify(formData), (err) => {
    if (err) throw err;
    console.log('The file has been saved!');
  });
  res.render('admin');
});

app.get("/userform",function(req,res){
  const n1labels = JSON.parse(fs.readFileSync('./uploads/'+selectedValue+'.docx.json'),'utf8')
  res.render('userform',{ n1labels });
});

const unzipper = require('unzipper');
const JSZip = require('jszip');
const Docxtemplater = require('docxtemplater');

// Load the Word template file


app.post("/userform",function(req,res){
  const data = JSON.parse(fs.readFileSync('./uploads/'+selectedValue+'.docx.json', 'utf8'));
  const formData = req.body;
  // Write the form data to a JSON file
  fs.writeFile('./uploads/'+selectedValue+'.docx.json', JSON.stringify(formData), (err) => {
    if (err) throw err;
    console.log('The file has been saved!');
    
});
 
});
  




let port = process.env.PORT;
if(port == null || port == "")
{
    port=3000;
}

app.listen(port,function(){
  console.log("Server started on port successfully");
});










