'use client';

import { useState } from 'react';
import * as XLSX from 'xlsx';

export default function ExcelConverter() {
  const [data, setData] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [keepCols, setKeepCols] = useState([]);
  const [renameMap, setRenameMap] = useState({});

  // Чтение и парсинг Excel файла
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    
    if (jsonData.length > 0) {
      setData(jsonData);
      const cols = Object.keys(jsonData[0]);
      setHeaders(cols);
      setKeepCols(cols);
      
      const initialRenameMap = {};
      cols.forEach(col => initialRenameMap[col] = col);
      setRenameMap(initialRenameMap);
    }
  };

  const toggleKeepCol = (col) => {
    setKeepCols(prev => 
      prev.includes(col) ? prev.filter(c => c !== col) : [...prev, col]
    );
  };

  const handleRenameChange = (oldName, newName) => {
    setRenameMap(prev => ({ ...prev, [oldName]: newName }));
  };

  // Генерация и скачивание CSV
  const handleExport = () => {
    if (data.length === 0) return;
    if (keepCols.length === 0) {
      alert("Выберите хотя бы одну колонку для экспорта!");
      return;
    }

    // 1. Формируем шапку (первую строку) с учетом переименований
    const finalHeaders = keepCols.map(col => renameMap[col] || col);
    const csvRows = [finalHeaders.join(';')];

    // 2. Формируем данные
    data.forEach(row => {
      const lineValues = keepCols.map(col => {
        let val = (row[col] || "").toString();
        // Жестко вычищаем переносы строк, чтобы не сломать структуру CSV
        return val.replace(/\r?\n|\r/g, '').trim(); 
      });

      // Соединяем значения точкой с запятой
      csvRows.push(lineValues.join(';'));
    });

    // 3. Собираем финальный текст с BOM-маркером для правильной кириллицы
    const csvString = '\uFEFF' + csvRows.join('\n');
    
    // 4. Скачиваем файл
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'processed_data.csv';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="max-w-3xl mx-auto p-6 space-y-8 font-sans">
      <h1 className="text-2xl font-bold text-gray-800">Конвертер Excel ➡️ CSV</h1>

      <div className="p-4 border-2 border-dashed border-gray-300 rounded-lg bg-white shadow-sm hover:bg-gray-50 transition-colors">
        <input 
          type="file" 
          accept=".xlsx, .xls" 
          onChange={handleFileUpload} 
          className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer"
        />
      </div>

      {headers.length > 0 && (
        <div className="space-y-6">
          <div className="space-y-4">
            <h2 className="text-xl font-semibold text-gray-700">Настройка экспорта</h2>
            <div className="border p-4 rounded-lg bg-gray-50 shadow-sm">
              <p className="text-sm text-gray-600 mb-4 font-medium">
                Снимите галочки с колонок, которые не нужны. В текстовых полях можно задать новые имена для шапки CSV.
              </p>
              
              <div className="max-h-[500px] overflow-y-auto space-y-3 pr-2">
                {headers.map(col => {
                  const isKept = keepCols.includes(col);
                  return (
                    <div key={col} className="flex items-center space-x-4 p-3 bg-white border rounded-md shadow-sm transition-opacity" style={{ opacity: isKept ? 1 : 0.6 }}>
                      <input 
                        type="checkbox" 
                        checked={isKept}
                        onChange={() => toggleKeepCol(col)}
                        className="rounded text-green-600 focus:ring-green-500 h-5 w-5 cursor-pointer"
                      />
                      <div className="flex-1 flex flex-col sm:flex-row sm:items-center sm:space-x-4">
                        <span className={`text-sm font-medium w-1/2 truncate ${isKept ? 'text-gray-700' : 'text-gray-400 line-through'}`} title={col}>
                          {col}
                        </span>
                        <input 
                          type="text" 
                          value={renameMap[col] || ''}
                          onChange={(e) => handleRenameChange(col, e.target.value)}
                          className="w-full sm:w-1/2 p-2 border border-gray-300 rounded text-sm focus:border-green-500 focus:ring-1 focus:ring-green-500 disabled:bg-gray-100 disabled:text-gray-400"
                          disabled={!isKept}
                          placeholder="Новое имя"
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          <button 
            onClick={handleExport}
            className="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-lg shadow hover:bg-blue-700 transition-colors focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2"
          >
            Сохранить как CSV
          </button>
        </div>
      )}
    </div>
  );
}