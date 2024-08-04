// src/App.js
import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import moment from 'moment';
import './App.css'; // Import the CSS file

function App() {
  const [data, setData] = useState([]);
  const [dateColumns, setDateColumns] = useState([]);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const binaryStr = e.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });

      // Find the sheet named 'User Details'
      const sheetName = workbook.SheetNames.find(name => name.toLowerCase() === 'user details');
      if (!sheetName) {
        alert('Sheet named "User Details" not found.');
        return;
      }
      
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      let headerRowIndex = -1;
      const dateColIndices = [];
      const headers = rows[0] || [];

      // Iterate through rows to find the header row containing 'date'
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (row.some((cell) => typeof cell === 'string' && cell.toLowerCase().includes('date'))) {
          headerRowIndex = i;
          row.forEach((cell, index) => {
            if (typeof cell === 'string' && cell.toLowerCase().includes('date')) {
              dateColIndices.push(index);
            }
          });
          break;
        }
      }

      if (headerRowIndex !== -1) {
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: headerRowIndex });

        // Convert serial date numbers and other date formats to MM/DD/YYYY
        const formattedData = jsonData.map((row, rowIndex) => {
          if (rowIndex === 0) return row; // Skip header row
          // Ensure each row has the same number of columns
          return headers.map((_, colIndex) => {
            const cell = row[colIndex];
            if (dateColIndices.includes(colIndex)) {
              if (typeof cell === 'number') {
                return formatExcelDate(cell);
              } else if (typeof cell === 'string') {
                return formatStringDate(cell);
              }
            }
            return cell !== undefined ? cell : ''; // Ensure empty cells are handled
          });
        });

        // Set data, ensuring headers are included only once
        setData([headers, ...formattedData.slice(1)]);
        setDateColumns(dateColIndices);

        // Log the selected column numbers containing 'date'
        console.log('Date columns indices:', dateColIndices);
      }
    };
    reader.readAsBinaryString(file);
  };

  const formatExcelDate = (serial) => {
    const utcDays = Math.floor(serial - 25569);
    const dateInfo = new Date(utcDays * 86400 * 1000);
    const year = dateInfo.getFullYear();
    const month = (`0${dateInfo.getMonth() + 1}`).slice(-2);
    const day = (`0${dateInfo.getDate()}`).slice(-2);
    return `${month}/${day}/${year}`;
  };

  const formatStringDate = (dateString) => {
    const parsedDate = moment(dateString, [
      'MM-DD-YYYY', 'DD-MM-YYYY', 'DD-MMM-YYYY', 'MMM-DD',
      'MMM DD, YYYY', 'DD MMM YYYY', 'MM/DD/YYYY', 'YYYY-MM-DD',
      'MMMM DD, YYYY', 'DD MMM', 'MMM-YYYY', 'MM-DD-YY',
      'M-D-YY', 'M-D-YYYY', 'MM-DD-YYYY HH:mm', 'MMMM D, YYYY',
      'MM-DD-YYYY h:mm A', 'MMMM D, YYYY h:mm A', 'dddd, MMMM D, YYYY'
    ], true);

    if (parsedDate.isValid()) {
      return parsedDate.format('MM/DD/YYYY');
    }
    return dateString; // Return original string if it can't be parsed
  };

  return (
    <div className="App">
      <h1>Excel File Reader</h1>
      <input type="file" onChange={handleFileUpload} />
      {data.length > 0 && (
        <table className="data-table">
          <thead>
            <tr>
              {data[0].map((key, index) => (
                <th key={index} className={dateColumns.includes(index) ? 'date-column' : ''}>
                  {key}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.slice(1).map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, cellIndex) => (
                  <td key={cellIndex}>
                    {cell !== null && cell !== undefined ? cell : ''}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

export default App;
