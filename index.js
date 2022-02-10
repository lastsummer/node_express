const express = require('express');
const multer  = require('multer');
const readExcel = require('./readExcel')
const wang = require('./wang')

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
app.get('/wang', function(req, res){
  res.render('wangLogin')
})
app.get('/wang/dashboard', async function(req, res){
  const data = await wang.getCurrentMonth()
  res.render('wangDashboad',{data})
})
app.get('/wang/search', async function(req, res){
  const data = await wang.getMonthData(req.query.year*1, req.query.month*1-1)
  res.json({data})
})
app.get('/wang/saveTime', async function(req, res){
  const data = await wang.saveTime(req.query.year*1, req.query.month*1, req.query.day, req.query.start, req.query.end)
  res.json({data})
})
app.get('/wang/deleteTime', async function(req, res){
  const data = await wang.deleteTime(req.query.year*1, req.query.month*1, req.query.day)
  res.json({data})
})

let port = 8080;
app.listen(port);