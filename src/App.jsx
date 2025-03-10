import React, { useState } from "react";
import * as XLSX from "xlsx";
import "bootstrap/dist/css/bootstrap.min.css";
import "./App.css";

function App() {
  const [excelData, setExcelData] = useState([]);
  const [showButton, setShowButton] = useState(false);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];

    if (!file) return;

    // file reader api
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workBook = XLSX.read(data, {
        type: "array",
      });
      const worksheet = workBook.Sheets[workBook.SheetNames[0]];

      const newData = XLSX.utils.sheet_to_json(worksheet);

      setExcelData((prevData) => [...prevData, ...newData]);
    };
    reader.readAsArrayBuffer(file);

    alert("file add");
    setShowButton(true);
  };

  const exportToExcel = () => {
    if (excelData.length === 0) {
      alert("No data to export!");
      return;
    }

    const workBook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(excelData);

    // ✅ Add Auto Filter (A1:A100 if data is long)
    worksheet["!autofilter"] = {
      ref: `A1:${String.fromCharCode(64 + Object.keys(excelData[0]).length)}${
        excelData.length + 1
      }`,
    };
    const colCount = Object.keys(excelData[0]).length;
    worksheet["!cols"] = Array(colCount).fill({ wch: 25 });

    XLSX.utils.book_append_sheet(workBook, worksheet, "Sheet1");

    XLSX.writeFile(workBook, "exported_data_with_filter.xlsx");
  };

  return (
    <div className="container mt-4">
      <h1 className="mb-4">Read Excel file in React</h1>
      <input
        type="file"
        accept=".xlsx"
        required
        className="form-control mb-4"
        onChange={handleFileUpload}
      />

      {showButton && (
        <button class="btn btn-primary" type="submit" onClick={exportToExcel}>
          Export to excel
        </button>
      )}
      {excelData && (
        <div>
          <h2>Excel Data</h2>
          <table className="table table-bordered">
            <thead>
              <tr>
                {Object.keys(excelData[0] || {}).map((key) => (
                  <th key={key}>{key}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {excelData.map((row, index) => (
                <tr key={index}>
                  {Object.values(row).map((value, i) => (
                    <td key={i}>{value}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default App;
