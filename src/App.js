import { useState } from "react";

import styles from './App.module.css';

const XLSX = window.XLSX;

function App() {
  // 上传的文件
  const [file, setFile] = useState(new ArrayBuffer());
  // 要生成的网页表格内容
  const [html, setHTML] = useState('');
  // 表格sheet列表
  const [sheets, setSheets] = useState([]);
  // 目标sheet
  const [defSheet, setDefSheet] = useState('');
  // 转换出来的JSON数据
  const [JSONData, setJSONData] = useState({});
  // 是否把工作表名称加入文件名
  const [addSheetName, setAddSheetName] = useState(false);

  // excel转换为网页表格函数
  const excelToTable = async (file, sheet) => {
    const table = await file.arrayBuffer();

    const wb = XLSX.read(table);
    const ws = wb.Sheets[sheet];

    setHTML(XLSX.utils.sheet_to_html(ws));
  };

  // excel转换为JSON函数
  const excelToJSON = (file) => {
    const fileReader = new FileReader();

    fileReader.onload = (e) => {
      try {
        const { result } = e.target;

        // 以二进制流方式读取得到整份excel表格对象
        const workbook = XLSX.read(result, { type: 'binary' });

        // 存储获取到的数据
        let data = {};

        // 遍历每张工作表进行读取
        for (const sheet in workbook.Sheets) {
          let tempData = [];

          if (workbook.Sheets.hasOwnProperty(sheet)) {
            // 利用 sheet_to_json 方法将 excel 转成 json 数据
            data[sheet] = tempData.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
          }
        }

        console.log(data);
        //将处理好的数据赋值给state
        setJSONData(data);
      } catch (e) {
        console.log(e);
        alert('文件类型不正确');
        return;
      };
    };

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(file);
  };

  // 添加下载链接函数
  const addLink = (data, file) => {
    const fileField = document.getElementById('file');

    // 如已有生成的下载链接，则先删除
    const prevLink = fileField.querySelector('a');
    if (prevLink) {
      prevLink.parentNode.removeChild(prevLink);
    }

    // 根据上传的文件名自动生成JSON名称
    let fileNameArray = file.name.split('.');
    let fileName = fileNameArray.slice(0, fileNameArray.length - 1);

    if (addSheetName) {
      fileName.push(defSheet);
    }

    fileName = fileName.join('.');

    // 根据文件生成下载的数据
    const blob = new Blob([JSON.stringify(data, null, 2)]);

    // 生成下载链接并插入网页
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${fileName}.json`;
    link.innerText = '下载JSON';
    fileField.append(link);
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

    setFile(files[0]);

    // 获取当前表格sheet列表
    const table = await files[0].arrayBuffer();
    const wb = XLSX.read(table);
    setSheets(wb.SheetNames);

    // 生成表格预览
    excelToTable(files[0], wb.SheetNames[0]);

    setDefSheet(wb.SheetNames[0]);

    // 将表格转换为JSON数据
    excelToJSON(files[0]);
  };

  const generateJSON = () => {
    // 添加下载链接
    addLink(JSONData[defSheet], file);
  }

  return (
    <>
      <fieldset>
        <legend>说明</legend>
        <a href="https://github.com/Phil-Libra/excel-to-json">源代码</a>
        <p>生成的文件名格式：源文件名.选择的工作表名(可选).json</p>
        <br />
        <p>仅支持工作表文件上传（含Excel及OpenDocument），其他文件转换会存在Bug。</p>
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
        <input type="file" name="excel-file" id="excel-file" onChange={handleUpload} />
        <input type="checkbox" onChange={() => setAddSheetName((prevState) => !prevState)} />把工作表名称加入文件名
        <button onClick={generateJSON}>生成JSON</button>
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