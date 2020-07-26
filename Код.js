// константы для формирования запроса к механизму авторизации VK
const vkAPIParamRedirectUri = "https://script.google.com/macros/s/AKfycbyZjOx007HPa2TdTWlteVXRb91s24seT5Aji_aSgnA/exec?mode=code"
const vkAPIParamClientId = 7549426
const vkAPIParamScope = 4+65536 // photos + offline, т.е. доступ к фотографиям и запрос на бессрочный токен
  
// константы для работы с листами таблицы
const optionsSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings")
const resultsSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("results")
  
// переменные для хранения токена доступа к интерфейсу API VK,  и хранения ссылок на выбираемые диапазоны данных из таблицы
var access_token = ""
var tmpRange = ""

function doGet(e) {
  // Logger.log(e)
  if(e.parameter['mode']=="start") {
    // первый шаг авторизации на VK - запрос кода
    var requestURI = "https://oauth.vk.com/authorize?client_id=" + vkAPIParamClientId + "&display=page&redirect_uri=" 
                     + vkAPIParamRedirectUri + "&response_type=code&v=5.120&scope=" + vkAPIParamScope
    return HtmlService.createHtmlOutput("<a target=\"new\" href=\"" + requestURI + "\">Получить Access token на VK.</a>")
  } else if (e.parameter['mode']=="code") {
    // второй шаг авторизации на VK - код получен от VK, запрос access_token, получение и запись в таблицу
    // code - параметр, который возвращает авторизационный скрипт VK
    // его надо передать от приложения напрямую для получения access_token
    const code = e.parameter['code']
    var requestURI = "https://oauth.vk.com/access_token?client_id=" + vkAPIParamClientId + "&client_secret=545CkJfuzxYhx6koJCd3&redirect_uri=" 
                     + vkAPIParamRedirectUri + "&code=" + code;
    // токен напрямую запрашивается скриптом
    var response = UrlFetchApp.fetch(requestURI);
    var responseParsed = JSON.parse(response); 
    tmpRange = optionsSht.getRange(1, 1, 1, 2)
    access_token = responseParsed.access_token
    
    // полученный access_token записывается в таблицу для последующего использования
    tmpRange.setValues([["access_token:",access_token]])
    
    return HtmlService.createHtmlOutput("access_token = " + access_token)
  } else if (e.parameter['mode']=="process") {
    return HtmlService.createHtmlOutput(process())
  } else
    return HtmlService.createHtmlOutput(0);
}

function process() {
    // получение object_id группы, указанной на странице опций в таблице
    // чтение access_token из таблицы
    access_token = optionsSht.getRange(1, 2).getValue()
    // чтение имени группы из таблицы
    const groupName = optionsSht.getRange(2, 2).getValue()
    // чтение числа обрабатываемых постов
    const postsToProcess = optionsSht.getRange(4, 2).getValue()
        
    // запрос object id для группы по имени группы
    var requestURI = "https://api.vk.com/method/utils.resolveScreenName?screen_name=" + groupName + "&access_token=" + access_token + "&v=5.120"
    var response = UrlFetchApp.fetch(requestURI);
    var responseParsed = JSON.parse(response);
    const groupId = responseParsed.response.object_id;    
    // запись полученного ID в таблицу
    optionsSht.getRange(3, 2).setValue(groupId)
    
    // запрос [последних по времени] 10 записей со стены группы по object id (object id для группы нужно передавать, умножив на -1)
    requestURI = "https://api.vk.com/method/wall.get?owner_id=" + (-1)*groupId + "&count=" + postsToProcess + "&access_token=" + access_token + "&v=5.120"
    response = UrlFetchApp.fetch(requestURI);
    responseParsed = JSON.parse(response);
    // сохраняем JSON объект списка постов в переменную
    var postsList = responseParsed.response.items 
    
    var i = 0
    var maxR = 0
    // формируем массив с данными для последующей записи в range таблицы
    var tmpArr = []
    // формируем строку заголовка
    tmpArr[0] = [["Id поста:"],["дата и время:"],["Подпись к картинке:"],["Картинка из поста (не ссылка):"]]

    // перебираем список постов
    for(var key in postsList) {
      // если объект attachments массив, то есть к посту есть вложения, и вложение в сообщение является фотографией (т.е. не видео, не ссылка и т.п.), 
      // то сохраняем в массив для последующего вывода в таблицу
      if(Array.isArray(postsList[key].attachments) && postsList[key].attachments[0].type == "photo") {
          tmpArr[i+1] = new Array(4)
          tmpArr[i+1][0]=postsList[key].id // id поста
          tmpArr[i+1][1]=Intl.DateTimeFormat('ru-RU', {timeStyle: "short",dateStyle: "short"}).format(postsList[key].date*1000) // форматированная дата поста
          tmpArr[i+1][2]=postsList[key].attachments[0].photo.text // подпись к картинке (не путать с текстом поста, т.е. к картинке есть своя подпись - не всегда)
          
          // определяем соотношение сторон изображения Ш/В
          var imageR = postsList[key].attachments[0].photo.sizes[0].width/postsList[key].attachments[0].photo.sizes[0].height
          // запоминаем максимальное значение соотношения сторон, чтобы потом установить максимальную ширину столбца с изображениями
          if(imageR > maxR) maxR = imageR
          var imageW = 100*imageR
          var imageH = 100
          tmpArr[i+1][3]="=IMAGE(\"" + postsList[key].attachments[0].photo.sizes[0].url + "\"; 4; " + imageH + "; " + Math.round(imageW) + ")"
          
          i++
        }
    }
    // очистка ранее загруженных данных (включая форматы и форматирование), если есть
    resultsSht.clear()
    // получение диапазона ячеек из таблицы в соответствии с количеством отобранных постов + строка заголовка (тут можно использовать, например, i+1)
    tmpRange = resultsSht.getRange(1, 1, tmpArr.length, 4)
    // запись подготовленного выше массива с данными в таблицу
    tmpRange.setValues(tmpArr)
    // получение диапазона ячеек столбца С из таблицы в соответствии с количеством отобранных постов + строка заголовка
    // для установки правила переноса данных в ячейке
    tmpRange = resultsSht.getRange(1, 3, tmpArr.length, 3)
    tmpRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    // установка высоты и ширины ячеек в столбце D
    resultsSht.setRowHeights(2, i, 100)
    resultsSht.setColumnWidths(4, 1, 100*maxR)
    
    return i;
}
