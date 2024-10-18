'use strict';  

const PROP_KEY_CURRENT_FOLDER = 'CurrentFolderInfo';    // 現在のフォルダ情報を保存するキー
const PROP_KEY_FILE_TABLE_CONDITION = 'FileTableCondition';   // 課題ファイル表示条件のキー
const propKeys = [PROP_KEY_CURRENT_FOLDER,PROP_KEY_FILE_TABLE_CONDITION];
const nameList = {}; // Classroom フォルダの採点のスプレッドシート"name_list"から読み込む名前リスト

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

function saveProperties(fileTableConditions)
{
  const userProp = PropertiesService.getUserProperties();    // User Properties を取得
  userProp.setProperty(PROP_KEY_FILE_TABLE_CONDITION,
		       JSON.stringify(fileTableConditions)); // 課題ファイル表示条件を保存
}

function makeSubmissionFileTable(fileId,fileTableConditions)
{
  saveProperties(fileTableConditions);
  getNameListFromParentFolder(fileId);
  return makeFileTable(makeFileList(fileId,fileTableConditions.tableSort,
				    fileTableConditions.submissionTarget,
				    fileTableConditions.getCommentsFlag),
		       fileTableConditions.imageWidth);
}

function getFileTableConditions(userProp)
{
  const fileTableCondProp =
	userProp.getProperty(PROP_KEY_FILE_TABLE_CONDITION); // 課題ファイル表示条件を取得
  if(fileTableCondProp){
    return JSON.parse(fileTableCondProp);
  } else {
    return false;
  }
}

function getNameListFromParentFolder(fileId)
{
  // fileId のフォルダの親フォルダにある採点のスプレッドシート "name_list" から
  // emailAddress をキーにした氏名リストを作る
  const getNameRe = /(.+)\((.+)\)/;                            // 名前取得用正規表現
  const file = DriveApp.getFileById(fileId);
  const parentFolders = file.getParents(); // ファイルの親フォルダを取得
  if(parentFolders.hasNext()) {
    const parentFolder = parentFolders.next();
    const folder = DriveApp.getFolderById(parentFolder.getId());
    const files = folder.getFilesByName('name_list');  // "name_list" を取得
    if (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) { // ファイルのMIMEを確認
        const spreadsheet = SpreadsheetApp.openById(file.getId());
        const sheet = spreadsheet.getSheets()[0];            // 1つ目のシートを取得
        const lastRow = sheet.getLastRow();                  // 最終行番号を取得
        const firstRow = 6;
        const rangeValues = sheet.getRange(firstRow, 1, lastRow-firstRow+1, 3).getValues();
        for (var i = 0; i < rangeValues.length; i++) {
          const matchName0 = getNameRe.exec(rangeValues[i][0]);
          const matchName1 = getNameRe.exec(rangeValues[i][1]);
          nameList[rangeValues[i][2]] = {nameJ:matchName0[1]+' '+matchName1[1],
                                         nameE:matchName0[2]+' '+matchName1[2]};
        }
      }
    }
  }
  
}

function makeFileList(fileId,tableSort,submissionTarget,getCommentsFlag)
{
  const timeZone = Session.getScriptTimeZone();         // 日時作成用にタイムゾーンを取得
  const userFiles = {};                 // ユーザーごとにファイルをまとめるための連想配列
  const folderContents = getFolderContents(fileId,getCommentsFlag);
  const weekJpTable = {Sun:'日', Mon:'月', Tue:'火', Wed:'水', Thu:'木', Fri:'金', Sat:'土'};
  folderContents.files.forEach(file => {
    if((submissionTarget != 'all' && file.submissionType == submissionTarget) ||
       (submissionTarget == 'all' && (file.submissionType == 'turned_in' ||
				      file.submissionType == 'returned'))){
      const email = file.sharingUser.emailAddress;
      let studentId = email;                      // 学生IDをを仮設定
      if(typeof getStudentId == 'function'){
	studentId = getStudentId(studentId);      // getStudentId 関数があったら学生IDを再設定
      }
      let studentName = file.sharingUser.displayName; // 学生名を仮設定
      if(file.sharingUser.emailAddress in nameList){
        studentName = (`${nameList[file.sharingUser.emailAddress].nameJ} `+
                       `(${nameList[file.sharingUser.emailAddress].nameE})`);
      } else if(typeof getStudentName == 'function'){
	studentName = getStudentName(studentName);    // getStudentName 関数があったら学生名を再設定
      }
      let departmentName = false;                  // 学部名を仮設定
      if(typeof getDepartmentName == 'function'){
	departmentName = getDepartmentName(email); // getDepartmentName 関数があったら学部名を設定
      }
      const student = {id : studentId, name: studentName, departmentName};
      const modifiedDate = new Date(file.modifiedDate);
      const weekJp = weekJpTable[Utilities.formatDate(modifiedDate, timeZone,'E')] || 'E';
      const datetime = Utilities.formatDate(modifiedDate, timeZone,  // ファイルの日時を取得
					    `yyyy/MM/dd (${weekJp}) HH:mm:ss`);
      // email をキーにしてデータを作成
      if(userFiles[email]){
	userFiles[email].files.push({file,lastUpdatedTime:datetime});
      } else {
	userFiles[email] = {user:file.sharingUser,student,files:[{file,lastUpdatedTime:datetime}]};
      }
    }
  });

  // 連想配列を基に配列を作る
  const fileList = [];
  Object.keys(userFiles).forEach(key => {  // キーでループ
    const elem = userFiles[key];
    if(elem.files.length > 1){             // ファイルが2個以上なら
      elem.files.sort((a,b)=>{             // 提出日時でソート
	if(a.lastUpdatedTime > b.lastUpdatedTime){
	  return 1;
	} else if(a.lastUpdatedTime < b.lastUpdatedTime){
	  return -1;
	} else {
	  return 0;
	}
      });
    }
    elem.submitTime = elem.files[elem.files.length-1].lastUpdatedTime; // 最後のファイルの日時
    fileList.push(elem);   // 配列に要素を追加
  });

  // ファイルリストのソート
  const sortFileList = (getSortKey, asc) => {
    fileList.sort((a,b)=>{
      if(getSortKey(a) > getSortKey(b)){
	return asc ?  1:-1;
      } else if(getSortKey(a) < getSortKey(b)){
	return asc ? -1: 1;
      } else {
	return 0;
      }
    });
  }

  if(tableSort.type == 'student_id'){
    sortFileList(key=>key.student.id,tableSort.style == 'asc' ? true:false);
  } else if(tableSort.type == 'submit_time'){
    sortFileList(key=>key.submitTime,tableSort.style == 'asc' ? true:false);
  }
//  sortFileList(key=>key.student.name,true);
//  sortFileList(key=>key.submitTime,true);
//  sortFileList(key=>key.submitTime,false);

  /*
  // 確認用
  fileList.forEach(elem => {
    Logger.log(`${elem.user.getEmail()} ${elem.user.getName()}`);
    elem.files.forEach(fileElem => {
      Logger.log(`file: ${fileElem.file.getName()} ${fileElem.lastUpdatedTime}`);
    });
  });
  */
  return fileList;
}

function getFileRole(file)
{
  // ファイル共有タイプの取得
  const roleTable = {owner:'オーナー', organizer:'管理者', fileOrganizer:'ファイル管理者',
		     writer:'編集者', reader:'閲覧者'};
  let role;
  if(file.userPermission){
    role = roleTable[file.userPermission.role];
  }
  return role;
}

function makeFileTable(fileList,imageWidth)
{
  const mimeRe = /(.+)\/(.+)/;                       // mime解析用正規表現
  let html = '';
  html += '<table class="list">';
  const listCnt = fileList.length;
  fileList.forEach((elem,index) => {
    let departmentName = '';
    if(elem.student.departmentName){
      departmentName = ', '+elem.student.departmentName;
    }
    html+=(`<tr><td>${index+1}/${listCnt}</td><td colspan="2">`+
	   `${elem.student.id} ${elem.student.name}, `+        // 学生ID,氏名
	   `${elem.user.emailAddress}${departmentName} `+      // email,学部名
 	   `[${elem.submitTime}]</td></tr>\n`);                // 最後のファイルの日時
    let fileCnt = elem.files.length;                           // ファイルの個数を設定
    elem.files.forEach((fileElem,fileIndex) => {               // ファイルのループ
      const file = fileElem.file;
      const matchMime = mimeRe.exec(file.mimeType);       // ファイルの種類を取得
      const role = getFileRole(file);                     // ファイル共有タイプの取得
      html += '<tr class="image_file">';
      html += (`<td></td><td valign="top">${fileIndex+1}/${fileCnt}</td>`+    // 個数
    	       `<td data-email="${file.sharingUser.emailAddress}">`+
	       `<span class="student-id">${elem.student.id}</span> `+         // 学生ID
	       `<span class="student-name">${elem.student.name}</span> `+     // 学生氏名
    	       `<span class="file-mime-type">[${matchMime[1]},${matchMime[2]}]</span>`);// mime type
      if(role){
	html += ` （ファイル共有:<span class="file-role">${role}</span>）`;  // ファイル共有タイプ
      }
      html += '<br>';
      html += (`file_name: <span class="file-name">${file.title}</span> `+     // ファイル名
	       `[<span class="file-datetime">${fileElem.lastUpdatedTime}</span>] `); // 日時
      if(!(matchMime[1] == 'image' && matchMime[2] != 'heif')){ // ファイルが表示可能画像でないとき
	html += `<img src="${file.iconLink}"> `;                // 目印用にアイコンを表示
      }
      html += `<a href="${file.alternateLink}" target="_blank">open</a> `; // open リンク
      // html += `<a href="${file.downloadUrl}">download</a><br>`;  // download リンク
      html += '<br>';
      if(file.comments && file.comments.length > 0){
	html += `コメント数: ${file.comments.length}`;
      }
      if(matchMime[1] == 'image' && matchMime[2] != 'heif'){   // ファイルが表示可能画像なら
      	html += (`<div class="image-canvas" id="file_id_${file.id}" `+ // 画像用の canvas 追加
      		 `data-reload-cnt="0" data-rot-angle="0" data-image-width="${imageWidth}"></div>`);
      } else if(matchMime[2] == 'pdf' || matchMime[2] == 'heif'){
	const imageHeight = imageWidth * 1.4141356; // A4 の比率
	html += (`<iframe class="embedded-file" style="display:none;" src="${file.embedLink}" `+
		 `width="${imageWidth}" height="${imageHeight}"></iframe>`);
      }
      html += '</td>';
      html += '</tr>';
    });
  });
  html += '</table>';
  return html;
}

function setSubmissionType(folderContents)
{
  // 提出状況情報の追加
  const folderOwnerEmailAddress = folderContents.folderInfo.owners[0].emailAddress
  folderContents.files.forEach(elem => {
    let type;
    if(elem.owners[0].emailAddress == folderOwnerEmailAddress){ // オーナーがフォルダのオーナと同じ
      if(elem.sharingUser){
	type = 'turned_in';  // 提出済み未返却
      } else {
	type = 'unknown';    // 不明
      }
    } else {
      if(elem.sharingUser){
	type = 'returned'; // 提出済み返却済み
      } else {
	type = 'disable';  // 削除された課題？
      }
    }
    elem.submissionType = type;
  });
}

function makeDriveDataList(prevId,fileId)
{
  const userProp = PropertiesService.getUserProperties(); // User Properties を取得
  if(prevId == '' && fileId == ''){
    const folderInfoProp =
	  JSON.parse(userProp.getProperty(PROP_KEY_CURRENT_FOLDER)); // 保存した値を所得
    if(!folderInfoProp){
      prevId = '';
      fileId = 'root';
    } else {
      prevId = '';
      fileId = folderInfoProp.fileId;
    }
  }
  if(prevId == null) prevId = '';
  userProp.setProperty(PROP_KEY_CURRENT_FOLDER,JSON.stringify({prevId,fileId}));
  const userEmailAddress = Drive.About.get().user.emailAddress; // 実行ユーザーのメールアドレス取得
  const folderContents = getFolderContents(fileId,false);  // fileId フォルダ以下の Drive のデータを取得
  const folderTitle = folderContents.folderInfo.title;
  const makeDataTable = (list,isFolder) => {  // table の作成
    let html = '';
    const type = isFolder ? 'folder':'file';
    if(isFolder && folderTitle != '' && folderTitle != null){
      html += `<h3 id="folder_title">${folderTitle}</h3>`;
    }
    html += `<table class="${type}_table">`;
    if(list.length > 0){
      html += ('<thead><tr><th></th><th></th><th>オーナー</th>'+
	       '<th colspan="2">最終更新</th></tr></thead>');
      html += '<tbody>';
      list.forEach((elem) => {
	let ownerName;                       // 所有者名
	if(elem.owners[0].emailAddress == userEmailAddress){
	  ownerName = '自分';                // 実行ユーザーの場合は「自分」と表示
	} else {
	  ownerName = elem.owners[0].displayName;
	}
	let lastModifyingUserName;           // 最終更新者名
	if(elem.lastModifyingUser){
	  if(elem.lastModifyingUser.emailAddress == userEmailAddress){
	    lastModifyingUserName = '自分';  // 実行ユーザーの場合は「自分」と表示
	  } else {
	    lastModifyingUserName = elem.lastModifyingUser.displayName;
	  }
	} else {
	  lastModifyingUserName = '';
	}
	html +=
	  `<tr class="${type} ${elem.submissionType}" data-file-id="${elem.id}">`+
	  `<td><img src="${elem.iconLink}"></td>`+
	  `<td class="file_title">${elem.title}</td>`+
	  `<td><span class="margin_left">${ownerName}</span></td>`+
	  `<td><span class="margin_left">${elem.modifiedDateTime}</span></td>`+
	  `<td><span class="margin_left">${lastModifyingUserName}</span></td>`+
	  '</tr>';
      });
      html += '</tbody>';
    }
    html += '</table>';
    return html;
  }
  const htmlData = {folderCtrlElems:''};
  if(folderContents.parentId){         // parentId が設定されていたら
    htmlData.folderCtrlElems +=   // 「上へ」を追加
      `<span class="folder_move" data-file-id="${folderContents.parentId}">上へ</span>`;
  } else {
    htmlData.folderCtrlElems += '<span class="folder_move disable">上へ</span>';
  }
  if(prevId != ''){               // top フォルダ以外
    htmlData.folderCtrlElems +=   // 「戻る」を追加
      `<span class="folder_move" data-file-id="${prevId}">戻る</span>`;
  } else {
    htmlData.folderCtrlElems += '<span class="folder_move disable" >戻る</span>';
  }
  htmlData.folder = makeDataTable(folderContents.folders,true); // フォルダ用の table 作成
  htmlData.file = makeDataTable(folderContents.files,false);    // ファイル用の table 作成
  htmlData.fileId = fileId;
  htmlData.fileTableConditions = getFileTableConditions(userProp);
  return htmlData;
}

function getParentIdList(id)
{
  // parents 情報から id を取得する
  const parents = Drive.Parents.list(id);
  const parentIdList = [];
  parents.items.forEach(elem => {
    parentIdList.push(elem.id);
  });
  return parentIdList;
}

function getFolderContents(id,getCommentsFlag)
{
  // id で指定したフォルダ内のファイル一覧をフォルダとそれ以外に分けて取得する
  // （id='root' のときはマイドライブ直下）
  const timeZone = Session.getScriptTimeZone();  // 日時作成用にタイムゾーンを取得
  const folderContents = {folderInfo:Drive.Files.get(id),    // 対象フォルダの情報
			  folders:[], files:[]}; // 対象フォルダ内のフォルダリストとファイルリスト
  const query = `"${id}" in parents and trashed = false`;
  let pageToken;
  do {
    const list = Drive.Files.list({
      q: query,
      maxResults: 100,
      pageToken: pageToken
    });
    if(list.items && list.items.length > 0) {
      for (var i = 0; i < list.items.length; i++) {
        const elem = list.items[i];
	const datetime = Utilities.formatDate(new Date(elem.modifiedDate), timeZone,
					      'yyyy/MM/dd HH:mm:ss');
	elem.modifiedDateTime = datetime;
	if(elem.mimeType == MimeType.FOLDER){  // フォルダを判定
	  folderContents.folders.push(elem);
	} else {
	  if(getCommentsFlag){
	    // ファイルのコメントを取得
	    elem.comments =
	      getPagedDataList(optionalArgs => Drive.Comments.list(elem.id,optionalArgs));
	  }
	  folderContents.files.push(elem);
	}
      }
    }
    pageToken = list.nextPageToken;
  } while (pageToken);

  const parentIdList = getParentIdList(id); // parentId のリストを取得
  if(parentIdList.length == 1){             // parentId がひとつなら
    folderContents.parentId = parentIdList[0];   // parentId を設定
  }

  setSubmissionType(folderContents);        // 提出状況情報の追加

  return folderContents;
}

function getPagedDataList(getList)
{
  // ページを進めながらデータを取得
  let nextPageToken = undefined;
  const list = [];
  do{
    const optionalArgs = {
      maxResults: 100
    };
    if(nextPageToken){
      optionalArgs.pageToken = nextPageToken;
    }
    const response = getList(optionalArgs);
    const elements = response.items;
    if(elements){
      list.push(...elements);
    }
    nextPageToken = response.nextPageToken;
  } while(nextPageToken != undefined);
  return list;
}

function deleteCurrentFolderInfoProperties()
{
  // 初期フォルダの設定を削除
  PropertiesService.getUserProperties().deleteProperty(PROP_KEY_CURRENT_FOLDER);
}

function deleteFileTableConditionProperties()
{
  // 課題一覧表示の設定を削除
  PropertiesService.getUserProperties().deleteProperty(PROP_KEY_FILE_TABLE_CONDITION);
}

function deleteAllUserProperties(key)
{
  // PropertiesService で使用している設定を全て削除
  propKeys.forEach(key => {
    PropertiesService.getUserProperties().deleteProperty(key)
  });
}

function getFileInfo(id)
{
  // ファイルの情報を取得
  return {file:Drive.Files.get(id),
	  permission:getPagedDataList(optionalArgs => Drive.Permissions.list(id,optionalArgs)),
	  revision:getPagedDataList(optionalArgs => Drive.Revisions.list(id,optionalArgs)),
	  // activity:JSON.stringify(DriveActivity.Activity.query({ancestorName: 'items/'+id}),null,4),
	  properties:getPagedDataList(optionalArgs => Drive.Properties.list(id,optionalArgs)),
	  comments:getPagedDataList(optionalArgs => Drive.Comments.list(id,optionalArgs))
	 };
}

function outputCsv(csvData,folderTitle)
{
  const timeZone = Session.getScriptTimeZone();         // 日時作成用にタイムゾーンを取得
  const datetime = Utilities.formatDate(new Date(), timeZone, 'yyyyMMdd_HHmm');
  const outputFileName = `${folderTitle}_${datetime}_files.csv`;
  DriveApp.createFile(outputFileName, csvData);
  return outputFileName;
}

function checkMimeTypeFolderNeq(type){
  return type != MimeType.FOLDER;    // フォルダでないことを判定
}

function checkMimeTypeFolderNeqAndMimeTypeEqPdf(type){
  return type != MimeType.FOLDER && type == MimeType.PDF;
}

function runCopy(targetFolderId,fileTableConditions,copySettings)
{
  saveProperties(fileTableConditions);
  getNameListFromParentFolder(targetFolderId);
  const folderInfo = Drive.Files.get(targetFolderId);   // コピー対象フォルダの情報取得
  const parentFolderInfo = Drive.Files.get(folderInfo.parents[0].id);
  let checkMimeTypeFunc = checkMimeTypeFolderNeq;
  if(copySettings.copy_file_type == 'PDF' || copySettings.copy_file_type == 'pdf'){
    checkMimeTypeFunc = checkMimeTypeFolderNeqAndMimeTypeEqPdf;
  }
  const files =                                         // コピー対象ファイル取得
	getFiles(`"${targetFolderId}" in parents and trashed = false`,
		 checkMimeTypeFunc);
  const outputFolder = makeOutputFolder(parentFolderInfo.title,
					folderInfo.title,
					copySettings.copy_target_folder_name); // 出力フォルダ作成
  const resultLog = copyFiles(files,folderInfo,outputFolder,copySettings,
			      fileTableConditions.submissionTarget,
                              checkMimeTypeFunc);     // コピー実行
  return {targetFolderName:folderInfo.title,
	  outputFolderName:outputFolder.getName(),result:resultLog};
}

function getFiles(query,checkMimeType)
{
  const files = [];
  let pageToken;
  do {
    const list = Drive.Files.list({
      q: query,
      maxResults: 100,
      pageToken: pageToken
    });
    if(list.items && list.items.length > 0) {
      for (var i = 0; i < list.items.length; i++) {
	if(checkMimeType(list.items[i].mimeType)){  // 種類を判定
          files.push(list.items[i]);
	}
      }
    }
    pageToken = list.nextPageToken;
  } while (pageToken);
  return files;
}

function makeFolder(folderName,parentFolder)
{
  // フォルダ名で指定したフォルダがなかったら作る
  const folderTemp = parentFolder.getFoldersByName(folderName);
  let folder;
  if(!folderTemp.hasNext()){
    folder = parentFolder.createFolder(folderName);   // なかったら作成
  } else {
    folder = folderTemp.next();                       // あったら取得
  }
  return folder;
}

function makeOutputFolder(targetClassrommFolderName,targetFolderName,copyTargetFolderName)
{
  // 出力フォルダの作成
  const rootFolder = makeFolder('Classroom_archives',DriveApp);             // root のフォルダ
  const classroomFolder = makeFolder(targetClassrommFolderName,rootFolder); // Classroom フォルダ
  const subFolder = makeFolder(copyTargetFolderName,classroomFolder);  // サブフォルダ
  const timeZone = Session.getScriptTimeZone();  // 日時作成用にタイムゾーンを取得
  const datetime = Utilities.formatDate(new Date(), timeZone,'yyyyMMdd');
  return makeFolder(`${targetFolderName}_copied_${datetime}`,subFolder);//コピー先フォルダ作成
}

function makeFileNameTableKey(elem)
{
  // copyFiles で使う fileNameTable で使うキーの作成
  return elem.sharingUser.emailAddress+elem.title;
}

const fileExtTable =   // 拡張子テーブル
      {'application/pdf':'.pdf',
       'application/zip':'.zip',
       'image/jpeg':'.jpg',
       'image/png':'.png',
       'image/heif':'.heic',
       'application/vnd.openxmlformats-officedocument.wordprocessingml.document':'docx',
       'application,vnd.openxmlformats-officedocument.spreadsheetml.sheet':'xlsx'};

function addFileExt(mimeType,fileName)
{
  // 拡張子テーブルを基にして、拡張しないファイルに拡張子を追加
  if(fileExtTable[mimeType]){
    const fileExt = fileExtTable[mimeType];
    if(fileName.substring(fileName.length-fileExt.length).toLowerCase() != fileExt){
      fileName += fileExt;
    }
  }
  return fileName;
}

function addFileCounter(mimeType,fileName,counter)
{
  // 重複ファイルに(1)(2)などを付けるときに、拡張子テーブルを基に拡張子の前に番号を追加
  if(fileExtTable[mimeType]){
    const fileExt = fileExtTable[mimeType];
    if(fileName.substring(fileName.length-fileExt.length).toLowerCase() != fileExt){
      fileName += `(${counter})`;
    } else {
      fileName = fileName.substring(0,fileName.length-fileExt.length)+`(${counter})`+fileExt;
    }
  }
  return fileName;
}

function setLatest(files)
{
  // 学生ごとの最新提出ファイルを見つける
  const fileUserTable = {};      // メールアドレスをキーにしたハッシュ
  for(let fileIdx = 0; fileIdx < files.length; fileIdx++){
    const elem = files[fileIdx];
    if(elem.sharingUser){        // 一度でも提出されたファイルを対象とする
      if(fileUserTable[elem.sharingUser.emailAddress]){
	fileUserTable[elem.sharingUser.emailAddress].push(elem);
      } else {
	fileUserTable[elem.sharingUser.emailAddress] = [elem];
      }
    }
  }
  Object.keys(fileUserTable).forEach(key => {  // キーでループ
    const elem = fileUserTable[key];
    if(elem.length > 1){             // ファイルが2個以上なら
      elem.sort((a,b)=>{             // 提出日時でソート（最新が先頭に来る）
	if(a.modifiedDate < b.modifiedDate){
	  return 1;
	} else if(a.modifiedDate > b.modifiedDate){
	  return -1;
	} else {
	  return 0;
	}
      });
    }
    elem[0].latest = true;  // 最新のものに latest を付ける
  });
}

function copyFiles(files,folderInfo,outputFolder,copySettings,submissionTarget,checkMimeTypeFunc)
{
  const folderOwnerEmailAddress = folderInfo.owners[0].emailAddress // 対象フォルダのオーナー取得
  const existFiles = getFiles(`"${outputFolder.getId()}" in parents and trashed = false`,
			      checkMimeTypeFunc);
  const existFileTable = {};           // 出力フォルダにあるファイルの名前をキーとした連想配列
  existFiles.forEach(elem => {
    existFileTable[elem.title] = true; // 出力フォルダにあるファイルの名前を保存
  });

  if(copySettings.copy_target == 'latest'){
    setLatest(files);                  // 最新提出ファイルの検出
  }

  // ファイルタイプの設定、対象ファイル数のカウント、各学生の重複ファイル数のカウント
  const fileNameTable = {};         // ファイル名をキーとした連想配列
  const submissionTargetType = {'turned_in':'ow_t'}; // UI の select の値と type の対応
  let targetFileCnt = 0;
  for(let fileIdx = 0; fileIdx < files.length; fileIdx++){
    const elem = files[fileIdx];
    let type;
    if(elem.owners[0].emailAddress == folderOwnerEmailAddress){ // オーナーがフォルダのオーナと同じ
      if(elem.sharingUser){
	type = 'ow_t';     // オーナーが自分, 提出済み未返却
      } else {
	type = 'unknown';  // 不明
      }
    } else {
      if(elem.sharingUser){
	type = 'ow_s';     // オーナーが自分以外, 提出済み返却済み?
      } else {
	type = 'disable';  // 削除された課題？
      }
    }
    elem.submissionType = type;
    elem.isTarget = false;
    if((submissionTarget != 'all' && elem.submissionType == submissionTargetType[submissionTarget]) ||
       (submissionTarget == 'all' && (elem.submissionType == 'ow_t' ||
				      elem.submissionType == 'ow_s'))){
      if((copySettings.copy_target != 'latest') ||
	 (copySettings.copy_target == 'latest' && elem.latest)){
	targetFileCnt++;
	elem.isTarget = true;;
	const fileNameTableKey = makeFileNameTableKey(elem);
	if(fileNameTable[fileNameTableKey]){  // 各学生で同じファイル名のファイルの数を数える
	  fileNameTable[fileNameTableKey]++;
	} else {
	  fileNameTable[fileNameTableKey] = 1;
	}
      }
    }
  }
  for(let key in fileNameTable){
    if(fileNameTable[key] == 1){    // ファイル数が1なら要素を削除
      delete fileNameTable[key];
    }
  }

  // コピーの実行
  let infoFileFolder;
  let copyFileCnt = 0;
  let skipFileCnt = 0;
  for(let fileIdx = 0; fileIdx < files.length; fileIdx++){
    const elem = files[fileIdx];
    if(elem.isTarget){
      const email = elem.sharingUser.emailAddress;
      let studentId = email;                      // 学生IDをを仮設定
      if(typeof getStudentId == 'function'){
	studentId = getStudentId(studentId);      // getStudentId 関数があったら学生IDを再設定
      }
      let studentName = elem.sharingUser.displayName;
      if(elem.sharingUser.emailAddress in nameList){
        studentName = nameList[elem.sharingUser.emailAddress].nameJ;
      } else if(typeof getName == 'function'){
	studentName = getName(studentName).nameJ; // 日本語名のみを使う
      }
//      const condition = (studentId == '20730047');
      const condition = true;
      if(condition){
	const file = DriveApp.getFileById(elem.id);
	let fileName = elem.title;
	const fileNameTableKey = makeFileNameTableKey(elem);
	if(fileNameTable[fileNameTableKey]){      // 各学生で同じファイル名のファイルがあったら
	  fileName = addFileCounter(file.getMimeType(),fileName,fileNameTable[fileNameTableKey]);
	  fileNameTable[fileNameTableKey]--;
	}
	// copySettings に応じて出ファイル名を設定する
	let outputFileName = '';
	if(copySettings.prefix != ''){
	  outputFileName += copySettings.prefix;
	}
	const addOutputFileName = (key,value)=>{
	  if(copySettings[key]){
	    if(outputFileName != '') outputFileName += '_';
	    outputFileName += value;
	  }
	}
	addOutputFileName('student_id',studentId);
	addOutputFileName('student_name',studentName);
	addOutputFileName('submission_type',elem.submissionType);
	addOutputFileName('orig_filename',fileName);
	if(copySettings.suffix != ''){
	  if(outputFileName != ''){
	    outputFileName += '_';
	  }
	  outputFileName += copySettings.suffix;
	}
	outputFileName = addFileExt(file.getMimeType(),outputFileName); // 拡張子の追加
	
	if(!existFileTable[outputFileName]){            // ファイルがなかったら
	  let activity, fileInfo;
	  try{
	    activity = DriveActivity.Activity.query({ancestorName: 'items/'+elem.id});
	    fileInfo = Drive.Files.get(elem.id);
	  } catch(error){
	    return {targetFileCnt, copyFileCnt, skipFileCnt, error};
	  }
	  if(!infoFileFolder){
	    infoFileFolder = makeFolder('info_files',outputFolder);
	  }
	  infoFileFolder.createFile(`${outputFileName}_info.json`, // activity, file を保存
				    JSON.stringify({activity,file:fileInfo}));
	  try{
	    file.makeCopy(outputFileName,outputFolder);   // ファイルをコピー
	  } catch(error){
	    return {targetFileCnt, copyFileCnt, skipFileCnt, error};
	  }
	  copyFileCnt++;
	} else {
	  skipFileCnt++;
	}
      }
    }
  }
  return {targetFileCnt, copyFileCnt, skipFileCnt, error:false};
}
