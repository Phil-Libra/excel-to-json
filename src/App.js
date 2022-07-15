import { useState } from "react";

import styles from './App.module.css';

const XLSX = window.XLSX;

function App() {
  // 上传的文件
  const [file, setFile] = useState();
  // 要生成的网页表格内容
  const [html, setHTML] = useState();
  // 表格sheet列表
  const [sheets, setSheets] = useState([]);
  // 目标sheet
  const [defSheet, setDefSheet] = useState();
  // 是否把工作表名称加入文件名
  const [addSheetName, setAddSheetName] = useState(false);
  // 设定下载按钮显示状态
  const [dlStatus, setDLStatus] = useState(false);

  // excel转换为网页表格函数
  const excelToTable = async (file, sheet) => {
    const table = await file.arrayBuffer();

    const wb = XLSX.read(table);
    const ws = wb.Sheets[sheet];

    setHTML(XLSX.utils.sheet_to_html(ws));
  };

  // excel转换为JSON函数
  const excelToJSON = (file) => new Promise((resolve, reject) => {
    const fileReader = new FileReader();

    fileReader.onload = (e) => {
      try {
        const { result } = e.target;

        // 以二进制流方式读取得到整份excel表格对象
        const workbook = XLSX.read(result, { type: 'binary' });

        // 存储获取到的数据
        const data = {};

        // 遍历每张工作表进行读取
        for (const sheet in workbook.Sheets) {
          const tempData = [];

          if (workbook.Sheets.hasOwnProperty(sheet)) {
            // 利用 sheet_to_json 方法将 excel 转成 json 数据
            data[sheet] = tempData.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
          }
        }

        // 返回处理好的数据
        resolve(data);
      } catch (e) {
        reject(e);
        return;
      };
    };

    // 以二进制方式打开文件
    fileReader.readAsArrayBuffer(file);
  });

  // 生成并下载文件函数
  const generateJSON = (data, file) => {
    // 根据上传的文件名自动生成JSON名称
    let fileNameArray = file.name.split('.');
    fileNameArray = fileNameArray.slice(0, fileNameArray.length - 1);

    // 如果上传的文件只有扩展名，则加入默认的文件名'json'
    if (!fileNameArray[0]) {
      fileNameArray.splice(0, 1, 'json');
    }

    if (addSheetName) {
      fileNameArray.push(defSheet);
    }

    const fileName = fileNameArray.join('.');

    // 根据文件生成下载的数据
    const blob = new Blob([JSON.stringify(data, null, 2)]);

    // 自动下载文件
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${fileName}.json`;
    link.click();
  };

  const handleUpload = async (e) => {
    // 获取文件并赋值给state
    const files = e.target.files;

    // 判定是否有文件上传及文件格式是否正确
    if (files.length === 0) {
      return;
    } else if (files.length !== 0 && !fileType.includes(files[0].type)) {
      alert('文件格式错误');
      return;
    }

    // 在state中共享文件
    setFile(files[0]);

    // 获取当前表格工作表列表
    const table = await files[0].arrayBuffer();
    const wb = XLSX.read(table);
    setSheets(wb.SheetNames);

    // 设置工作表
    setDefSheet(wb.SheetNames[0]);

    // 生成表格预览
    excelToTable(files[0], wb.SheetNames[0]);

    // 显示下载按钮
    setDLStatus(true);
  };

  const handleDownload = async () => {
    // // 考虑未上传任何文件的情况
    if (!file) {
      return;
    }

    // // 将表格转换为JSON数据
    const JSONData = await excelToJSON(file);
    // // 生成文件并下载
    generateJSON(JSONData[defSheet], file);
  };

  return (
    <>
      <fieldset>
        <legend>说明</legend>
        <p><a href="https://github.com/Phil-Libra/excel-to-json">源代码</a></p>
        <br />
        <p>生成的文件名格式：源文件名.选择的工作表名(可选).json</p>
        <br />
        <p>仅支持工作表文件上传（含Excel及OpenDocument），其他文件转换会存在Bug。</p>
        <br />
        <p><strong>没有任何值的行、列虽然在预览中能看到，但在转换时会被忽略。</strong></p>
        <br />
        <p>暂时仅支持如下格式表格转换，否则转换出的数据可能有bug：</p>
        <table>
          <tbody>
            <tr key="1">
              <td>JSON key1</td>
              <td>JSON key2</td>
              <td>JSON key3</td>
              <td>JSON key4</td>
              <td>JSON key5</td>
            </tr>
            <tr key="2">
              <td>key1 value</td>
              <td>key2 value</td>
              <td>key3 value</td>
              <td>key4 value</td>
              <td>key5 value</td>
            </tr>
            <tr key="3">
              <td>key1 value</td>
              <td>key2 value</td>
              <td>key3 value</td>
              <td>key4 value</td>
              <td>key5 value</td>
            </tr>
            <tr key="4">
              <td>key1 value</td>
              <td>key2 value</td>
              <td>key3 value</td>
              <td>key4 value</td>
              <td>key5 value</td>
            </tr>
            <tr key="5">
              <td>key1 value</td>
              <td>key2 value</td>
              <td>key3 value</td>
              <td>key4 value</td>
              <td>key5 value</td>
            </tr>
          </tbody>
        </table>
      </fieldset>

      <fieldset id='file'>
        <legend>上传文件</legend>
        <div>
          <label htmlFor="excel-file">
            <input
              type="file"
              name="excel-file"
              id="excel-file"
              onChange={handleUpload}
            />
          </label>
        </div>
        <div style={{ display: dlStatus ? '' : 'none' }}>
          <label htmlFor="addSheet">
            <input
              type="checkbox"
              id="addSheet"
              onChange={() => setAddSheetName((prevState) => !prevState)}
            />
            把工作表名称加入文件名
          </label>
          <button onClick={handleDownload}>下载JSON</button>
        </div>
      </fieldset>

      <fieldset>
        <legend>表格数据预览</legend>
        {
          sheets.length > 0
            ? (
              <>
                选择工作表：
                <select
                  id='sheets'
                  value={defSheet}
                  onChange={(e) => {
                    excelToTable(file, e.target.value);
                    setDefSheet(e.target.value)
                  }}
                >
                  {
                    sheets.map((item, index) => (
                      <option value={item} key={index}>{item}</option>
                    ))
                  }
                </select>

                <div
                  className={styles.tablePreview}
                  dangerouslySetInnerHTML={{ __html: html }}
                />
              </>
            )
            : (<></>)
        }
      </fieldset>
    </>
  )
};

export default App;

const fileType = [
  'application/vnd.ms-excel',
  'application/vnd.ms-excel.addin.macroEnabled.12',
  'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
  'application/vnd.ms-excel.sheet.macroEnabled.12',
  'application/vnd.ms-excel.template.macroEnabled.12',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
  'application/vnd.oasis.opendocument.spreadsheet'
];