const express = require('express');
const multer = require('multer');
const bodyParser = require('body-parser');
const readExcel = require('./readExcel')
const idExcel = require('./service/idExcel')
const pressExcel = require('./service/pressExcel')
const pressExcelNew = require('./service/pressExcelNew')
const sharpImage = require('./service/sharpImage')
const wang = require('./wang')

let app = express();
app.use(bodyParser.urlencoded({ extended: false }))
app.use(bodyParser.json())

app.use(express.static('public'));
app.set('view engine','ejs')
const upload = multer({ dest: 'uploads/' })


app.get('/', function (req, res) {
    res.send('Hello')
})

// ------------------------------------------------ 上傳excel 頁面
// 血壓 + 心力評量表
app.get('/upload', function(req, res){
  res.render('upload')
})
// 員工資料
app.get('/upload/id', function(req, res){
  res.render('uploadID')
})
// 員工資料(月份)
app.get('/upload/monthID/:month', function(req, res){
  res.render('uploadMonthID',{month:req.params.month})
})
// 血壓(月份)
app.get('/upload/press/:month', function(req, res){
  res.render('uploadMonthPress',{month:req.params.month})
})
// 離職員工資料
app.get('/upload/removeId', function(req, res){
  res.render('uploadremoveId')
})

// ------------------------------------------------ 上傳excel api
// 心力評量表
app.post('/import/Upload4', upload.single('uploadExcel4'), async function(req, res){
  console.log(req.file);
  const fileName = await readExcel.parserExcel(req.file.path)
  res.json({fileName})
})
// 離職員工資料
app.post('/import/UploadRemoveId', upload.single('uploadExcelRemoveId'), async function(req, res){
  await idExcel.parserRemoveIDExcel(req.file.path)
  res.json("ok")
})
// 員工資料
app.post('/import/UploadID', upload.single('uploadExcelID'), async function(req, res){
  await idExcel.parserIDExcel(req.file.path)
  res.json("ok")
})
// 員工資料(月份)
app.post('/import/UploadMonthID/:month', upload.single('uploadExcelMonthID'), async function(req, res){
  console.log(req.file)
  await idExcel.parserMonthIDExcel(req.file.path, req.params.month)
  res.json("ok")
})
// 血壓
app.post('/import/UploadPress', upload.single('uploadExcelPress'), async function(req, res){
  console.log(req.file);
  const fileName = await pressExcel.parserPressExcel(req.file.path, '')
  res.json({fileName})
})
// 血壓Version
app.post('/import/UploadPressNew', upload.single('uploadExcelPressNew'), async function(req, res){
  console.log(req.file);
  const fileName = await pressExcelNew.parserPressExcel(req.file.path)
  res.json({fileName})
})

// 血壓(月份)
app.post('/import/UploadMonthPress/:month', upload.single('uploadExcelMonthPress'), async function(req, res){
  console.log(req.file);
  const fileName = await pressExcel.parserPressExcel(req.file.path, req.params.month)
  res.json({fileName})
})

// 下載excel / jpg
app.get('/download/:filename', function(req, res){
  const file = `./result/${req.params.filename}`;
  res.download(file);
});

// ------------------------------- 王俊皓

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

// ------------------------------------------------ messgae
app.get('/messgae', function(req, res){
  res.render('messgae')
})
app.post('/message/MessageCode', async function(req, res){
  // const fileName = "123"
  const fileName = await sharpImage.getMetadata(req.body.localCode)
  res.json({fileName})
})


let port = 80;
app.listen(port);