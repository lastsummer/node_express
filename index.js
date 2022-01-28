const express = require('express');
const multer  = require('multer');
const readExcel = require('./readExcel')

let app = express();
app.use(express.static('public'));
app.set('view engine','ejs')
const upload = multer({ dest: 'uploads/' })


app.get('/', function (req, res) {
    res.send('Hello')
})
app.get('/upload', function(req, res){
  res.render('upload')
})
app.post('/import/Upload4', upload.single('uploadExcel4'), async function(req, res){
  console.log(req.file);
  const fileName = await readExcel.parserExcel(req.file.path)
  res.json({fileName})
})
app.get('/download/:filename', function(req, res){
  const file = `./result/${req.params.filename}`;
  res.download(file);
});

let port = 8080;
app.listen(port);