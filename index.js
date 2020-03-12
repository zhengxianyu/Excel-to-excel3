$(function() {
  $('#complianceIndicators').change(function(parentEvent) {
    let files = parentEvent.target.files;

    let fileReader = new FileReader();

    let getExcelList = [];
    fileReader.onload = function(childEvent) {
      let excelBinaryData;
      // 读取上传的excel文件
      try {
        let excelData = childEvent.target.result;
        excelBinaryData = XLSX.read(excelData, {
          type: 'binary'
        });
      } catch (parentEvent) {
        console.log('该文件类型不能识别');
        return;
      }

      // 获取excel所有元素
      for (let sheet in excelBinaryData.Sheets) {
        if (excelBinaryData.Sheets.hasOwnProperty(sheet)) {
          let excelSheet = XLSX.utils.sheet_to_json(excelBinaryData.Sheets[sheet]);
          getExcelList[getExcelList.length] = excelSheet;
        }
      }
      // console.log("==getExcelList::")
      // console.log(getExcelList)

      let newExcelList = [];
      let roleWeightOne = {};
      let roleWeightTwo = {};
      let roleWeightThree = {};
      let roleWeightFour = {};
      let roalOne = [];
      let roalTwo = [];
      let roalThree = [];
      let roalFour = [];

      // 找出可以合并的数据
      for (let i = 0; i < getExcelList[0].length; i++) {
        if (getExcelList[0][i]['适用角色'] == '员工') {
          roleWeightOne[getExcelList[0][i]['指标名称']] = getExcelList[0][i]['指标权重'];
          roalOne.push(getExcelList[0][i]['指标名称']);
        } else if (getExcelList[0][i]['适用角色'] == '内控联系人') {
          roleWeightTwo[getExcelList[0][i]['指标名称']] = getExcelList[0][i]['指标权重'];
          roalTwo.push(getExcelList[0][i]['指标名称']);
        } else if (getExcelList[0][i]['适用角色'] == '专职合规岗') {
          roleWeightThree[getExcelList[0][i]['指标名称']] = getExcelList[0][i]['指标权重'];
          roalThree.push(getExcelList[0][i]['指标名称']);
        } else if (getExcelList[0][i]['适用角色'] == '部门负责人') {
          roleWeightFour[getExcelList[0][i]['指标名称']] = getExcelList[0][i]['指标权重'];
          roalFour.push(getExcelList[0][i]['指标名称']);
        }
      }

      for (let i = 0; i < getExcelList[1].length; i++) {
        for (let j = 0; j < roalOne.length; j++) {
          newExcelList[newExcelList.length] = getChildColumn(getExcelList[1][i], roleWeightOne, roalOne[j]);
        }
      }

      for (let i = 0; i < getExcelList[2].length; i++) {
        for (let j = 0; j < roalTwo.length; j++) {
          newExcelList[newExcelList.length] = getChildColumn(getExcelList[2][i], roleWeightTwo, roalTwo[j]);
        }
      }

      for (let i = 0; i < getExcelList[3].length; i++) {
        for (let j = 0; j < roalThree.length; j++) {
          newExcelList[newExcelList.length] = getChildColumn(getExcelList[3][i], roleWeightThree, roalThree[j]);
        }
      }

      for (let i = 0; i < getExcelList[4].length; i++) {
        for (let j = 0; j < roalFour.length; j++) {
          newExcelList[newExcelList.length] = getChildColumn(getExcelList[4][i], roleWeightFour, roalFour[j]);
        }
      }

      downloadFiveSheet(getExcelList, newExcelList, files);
    };

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
  });

  $('#contractData').change(function(parentEvent) {
    let files = parentEvent.target.files;

    let fileReader = new FileReader();

    let getExcelList = [];
    fileReader.onload = function(childEvent) {
      let excelBinaryData;
      // 读取上传的excel文件
      try {
        let excelData = childEvent.target.result;
        excelBinaryData = XLSX.read(excelData, {
          type: 'binary'
        });
      } catch (parentEvent) {
        console.log('该文件类型不能识别');
        return;
      }

      // 获取excel所有元素
      for (let sheet in excelBinaryData.Sheets) {
        if (excelBinaryData.Sheets.hasOwnProperty(sheet)) {
          let excelSheet = XLSX.utils.sheet_to_json(excelBinaryData.Sheets[sheet]);
          getExcelList[getExcelList.length] = excelSheet;
        }
      }

      let newExcelList = [];
      
      for (let i = 1; i < getExcelList[0].length; i++) {
        let item = getExcelList[0][i];
        let noOtherData = true;

        if (('承揽分配' in item) && ('__EMPTY' in item)) {
          let newItem = getChildItem(item, '承揽分配', '__EMPTY');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_1' in item) && ('__EMPTY_2' in item)) {
          let newItem = getChildItem(item, '__EMPTY_1', '__EMPTY_2');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_3' in item) && ('__EMPTY_4' in item)) {
          let newItem = getChildItem(item, '__EMPTY_3', '__EMPTY_4');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_5' in item) && ('__EMPTY_6' in item)) {
          let newItem = getChildItem(item, '__EMPTY_5', '__EMPTY_6');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_7' in item) && ('__EMPTY_8' in item)) {
          let newItem = getChildItem(item, '__EMPTY_7', '__EMPTY_8');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (noOtherData) {
          let newItem = getChildItem(item);
          newExcelList[newExcelList.length] = newItem;
        }
        
      }

      downloadOneSheet(newExcelList, files)
    };

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
  });

  $('#bearingData').change(function(parentEvent) {
    let files = parentEvent.target.files;

    let fileReader = new FileReader();

    let getExcelList = [];
    fileReader.onload = function(childEvent) {
      let excelBinaryData;
      // 读取上传的excel文件
      try {
        let excelData = childEvent.target.result;
        excelBinaryData = XLSX.read(excelData, {
          type: 'binary'
        });
      } catch (parentEvent) {
        console.log('该文件类型不能识别');
        return;
      }

      // 获取excel所有元素
      for (let sheet in excelBinaryData.Sheets) {
        if (excelBinaryData.Sheets.hasOwnProperty(sheet)) {
          let excelSheet = XLSX.utils.sheet_to_json(excelBinaryData.Sheets[sheet]);
          getExcelList[getExcelList.length] = excelSheet;
        }
      }

      let newExcelList = [];
      
      for (let i = 1; i < getExcelList[0].length; i++) {
        let item = getExcelList[0][i];
        let noOtherData = true;

        if (('承做负责' in item) && ('__EMPTY' in item)) {
          let newItem = getChildRoleItem(item, '承做负责', '承做负责', '__EMPTY');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_1' in item) && ('__EMPTY_2' in item)) {
          let newItem = getChildRoleItem(item, '承做负责', '__EMPTY_1', '__EMPTY_2');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('承做参与' in item) && ('__EMPTY_3' in item)) {
          let newItem = getChildRoleItem(item, '承做参与', '承做参与', '__EMPTY_3');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_4' in item) && ('__EMPTY_5' in item)) {
          let newItem = getChildRoleItem(item, '承做参与', '__EMPTY_4', '__EMPTY_5');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_6' in item) && ('__EMPTY_7' in item)) {
          let newItem = getChildRoleItem(item, '承做参与', '__EMPTY_6', '__EMPTY_7');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_8' in item) && ('__EMPTY_9' in item)) {
          let newItem = getChildRoleItem(item, '承做参与', '__EMPTY_8', '__EMPTY_9');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_10' in item) && ('__EMPTY_11' in item)) {
          let newItem = getChildRoleItem(item, '承做参与', '__EMPTY_10', '__EMPTY_11');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_12' in item) && ('__EMPTY_13' in item)) {
          let newItem = getChildRoleItem(item, '承做参与', '__EMPTY_12', '__EMPTY_13');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_14' in item) && ('__EMPTY_15' in item)) {
          let newItem = getChildRoleItem(item, '承做参与', '__EMPTY_14', '__EMPTY_15');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_16' in item) && ('__EMPTY_17' in item)) {
          let newItem = getChildRoleItem(item, '承做参与', '__EMPTY_16', '__EMPTY_17');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (('__EMPTY_18' in item) && ('__EMPTY_19' in item)) {
          let newItem = getChildRoleItem(item, '承做参与', '__EMPTY_18', '__EMPTY_19');
          newExcelList[newExcelList.length] = newItem;
          noOtherData = false;
        }

        if (noOtherData) {
          let newItem = getChildRoleItem(item, '');
          newExcelList[newExcelList.length] = newItem;
        }
        
      }

      downloadOneSheet(newExcelList, files);
    };

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
  });

  $('#wellStreet').change(function(parentEvent) {
    let files = parentEvent.target.files;

    let fileReader = new FileReader();

    let getExcelList = [];
    fileReader.onload = function(childEvent) {
      let excelBinaryData;
      // 读取上传的excel文件
      try {
        let excelData = childEvent.target.result;
        excelBinaryData = XLSX.read(excelData, {
          type: 'binary'
        });
      } catch (parentEvent) {
        console.log('该文件类型不能识别');
        return;
      }

      // 获取excel所有元素
      for (let sheet in excelBinaryData.Sheets) {
        if (excelBinaryData.Sheets.hasOwnProperty(sheet)) {
          let excelSheet = XLSX.utils.sheet_to_json(excelBinaryData.Sheets[sheet]);
          getExcelList[getExcelList.length] = excelSheet;
        }
      }

      let newExcelList = [];
      for (let i = 1; i < getExcelList[0].length; i++) {
        let keyList = Object.getOwnPropertyNames(getExcelList[0][i]);
        for (let j = 1; j < keyList.length; j++) {
          if (keyList[j].indexOf('课程') != -1) {
            let item = getAllCourse(keyList[j], getExcelList[0][i]);
            newExcelList[newExcelList.length] = item;
          }
        }
      }

      downloadOneSheet(newExcelList, files);
    };

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
  });

  function getAllCourse(keyName, oldItem) {
    let item = {};
    item['1、您的姓名：'] = oldItem['1、您的姓名：'];
    item['电子邮箱'] = oldItem['电子邮箱'];
    item['一级部门'] = oldItem['一级部门'];
    item['课程名称'] = oldItem[keyName];

    return item;
  }

  function getChildRoleItem(oldItem, role, keyOne, keyTwo) {
    let newItem = {};
    newItem['来源部门'] = oldItem['来源部门'];
    newItem['部门'] = oldItem['部门'];
    newItem['IBD项目编号'] = oldItem['IBD项目编号'];
    newItem['项目名称（简称）'] = oldItem['项目名称（简称）'];
    newItem['项目类型'] = oldItem['项目类型'];
    newItem['年份'] = oldItem['年份'];
    newItem['项目状态'] = oldItem['项目状态'];
    newItem['承做计分周期选择'] = oldItem['承做计分周期选择'];
    newItem['项目收入/万元'] = oldItem[' 项目收入/万元 '];
    newItem['角色'] = role;

    if (keyOne) {
      newItem['姓名'] = oldItem[keyOne];
    }
    if (keyTwo) {
      newItem['比例'] = oldItem[keyTwo];
    }

    return newItem;
  }

  function getChildItem(oldItem, keyOne, keyTwo) {
    let newItem = {};
    newItem['来源部门'] = oldItem['来源部门'];
    newItem['部门'] = oldItem['部门'];
    newItem['IBD项目编号'] = oldItem['IBD项目编号'];
    newItem['项目名称（简称）'] = oldItem['项目名称（简称）'];
    newItem['项目类型'] = oldItem['项目类型'];
    newItem['年份'] = oldItem['年份'];
    newItem['项目状态'] = oldItem['项目状态'];
    newItem['承做计分周期选择'] = oldItem['承做计分周期选择'];
    newItem['项目收入/万元'] = oldItem[' 项目收入/万元 '];

    if (keyOne) {
      newItem['承揽人'] = oldItem[keyOne];
    }
    if (keyTwo) {
      newItem['比例'] = oldItem[keyTwo];
    }

    return newItem;
  }

  function getChildColumn(columnItem, roleWeight, addColumnName) {
    let type = 0;
    let standard = '';
    let finishStatus = '';
    let startTime = '01/01/2019';
    let endTime = '12/31/2019';

    let columnTemp = {};
    columnTemp['用户名'] = columnItem['用户名'];
    columnTemp['类型'] = type;
    columnTemp['指标名称'] = addColumnName;
    columnTemp['指标权重'] = roleWeight[addColumnName];
    columnTemp['衡量标准'] = standard;
    columnTemp['开始时间'] = startTime;
    columnTemp['结束时间'] = endTime;
    columnTemp['完成情况'] = finishStatus;
    return columnTemp;
  }

  function downloadOneSheet(newExcelList, files) {
    const newSheet = {
      SheetNames: ['Sheet1'],
      Sheets: {},
      Props: {}
    };
    const sheetDownloadType = { bookType: 'xlsx', bookSST: false, type: 'binary' };

    newSheet.Sheets['Sheet1'] = XLSX.utils.json_to_sheet(newExcelList);
    saveAs(
      new Blob(
        [
          stringToArrayBuffer(XLSX.write(newSheet, sheetDownloadType))
        ], {
          type: "application/octet-stream"
        }
      ),
      files[0].name
    );
  }

  function downloadFiveSheet(oldData, newExcelList, files) {
    const newSheet = {
      SheetNames: ['Sheet0', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5'],
      Sheets: {},
      Props: {}
    };
    const sheetDownloadType = { bookType: 'xlsx', bookSST: false, type: 'binary' };

    newSheet.Sheets['Sheet0'] = XLSX.utils.json_to_sheet(oldData[0]);
    newSheet.Sheets['Sheet1'] = XLSX.utils.json_to_sheet(oldData[1]);
    newSheet.Sheets['Sheet2'] = XLSX.utils.json_to_sheet(oldData[2]);
    newSheet.Sheets['Sheet3'] = XLSX.utils.json_to_sheet(oldData[3]);
    newSheet.Sheets['Sheet4'] = XLSX.utils.json_to_sheet(oldData[4]);
    newSheet.Sheets['Sheet5'] = XLSX.utils.json_to_sheet(newExcelList);
    saveAs(
      new Blob(
        [
          stringToArrayBuffer(XLSX.write(newSheet, sheetDownloadType))
        ], {
          type: "application/octet-stream"
        }
      ),
      files[0].name
    );
  }

  function stringToArrayBuffer(data) {
    if (typeof ArrayBuffer !== 'undefined') {
      let arrayBuffer = new ArrayBuffer(data.length);
      let unitArray = new Uint8Array(arrayBuffer);
      for (let unitI = 0; unitI != data.length; unitI++) {
        unitArray[unitI] = data.charCodeAt(unitI) & 0xFF;
      }
      return arrayBuffer;
    } else {
      let arrayBuffer = new Array(data.length);
      for (let bufferI = 0; bufferI != data.length; bufferI++) {
        arrayBuffer[bufferI] = data.charCodeAt(bufferI) & 0xFF;
      }
      return arrayBuffer;
    }
  }

  function saveAs(content, fileName) {
    let clickDiv = document.createElement("a");
    clickDiv.download = fileName || "下载";
    clickDiv.href = URL.createObjectURL(content);
    clickDiv.click();
    setTimeout(function () {
      URL.revokeObjectURL(content);
    }, 100);
  }
})
