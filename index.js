let express = require('express');
  //1.載入express模組
let app = express();
  // 2.使用express
app.get('/', function (req, res) {
    res.send('Hello')
})
  // 5.首頁
let port = 80;
  //3.設定port位置
app.listen(port);