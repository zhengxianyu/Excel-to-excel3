$(function() {
  $('#excelFile').change(function(parentEvent) {
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
      console.log("==getExcelList::")
      console.log(getExcelList)

      let newExcelList = [];
      let roleWeight = {};
      let roalOne = [];
      let roalTwo = [];
      let roalThree = [];
      let roalFour = [];

      // 找出可以合并的数据
      for (let i = 0; i < getExcelList[0].length; i++) {
        roleWeight[getExcelList[0][i]['指标名称']] = getExcelList[0][i]['指标权重'];
        if (i < 4) {
          roalOne.push(getExcelList[0][i]['指标名称']);
        } else if (i >= 4 && i < 12) {
          roalTwo.push(getExcelList[0][i]['指标名称']);
        } else if (i >= 12 && i < 24) {
          roalThree.push(getExcelList[0][i]['指标名称']);
        } else if (i >=24 && i < 28) {
          roalFour.push(getExcelList[0][i]['指标名称']);
        }
      }

      for (let i = 0; i < getExcelList[1].length; i++) {
        for (let j = 0; j < roalOne.length; j++) {
          newExcelList[newExcelList.length] = getChildColumn(getExcelList[1][i], roleWeight, roalOne[j]);
        }
      }

      for (let i = 0; i < getExcelList[2].length; i++) {
        for (let j = 0; j < roalTwo.length; j++) {
          newExcelList[newExcelList.length] = getChildColumn(getExcelList[2][i], roleWeight, roalTwo[j]);
        }
      }

      for (let i = 0; i < getExcelList[3].length; i++) {
        for (let j = 0; j < roalThree.length; j++) {
          newExcelList[newExcelList.length] = getChildColumn(getExcelList[3][i], roleWeight, roalThree[j]);
        }
      }

      for (let i = 0; i < getExcelList[4].length; i++) {
        for (let j = 0; j < roalFour.length; j++) {
          newExcelList[newExcelList.length] = getChildColumn(getExcelList[4][i], roleWeight, roalFour[j]);
        }
      }

      download(getExcelList, newExcelList, files);
    };

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
  });

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

  function download(oldData, newExcelList, files) {
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
