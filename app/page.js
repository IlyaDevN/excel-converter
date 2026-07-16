'use client';

import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';

// ==========================================
// КОМПОНЕНТ 1: Excel -> CSV
// ==========================================
function ExcelToCsv() {
  const [data, setData] = useState([]);
  const [orderedCols, setOrderedCols] = useState([]);
  const [keepCols, setKeepCols] = useState([]);
  const [renameMap, setRenameMap] = useState({});
  const [csvSeparator, setCsvSeparator] = useState(';');
  const [dateFormat, setDateFormat] = useState('original');

  const dragItem = useRef(null);
  const dragOverItem = useRef(null);

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { cellDates: true });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
      defval: "", 
      raw: false 
    });
    
    if (jsonData.length > 0) {
      setData(jsonData);
      const cols = Object.keys(jsonData[0]);
      setOrderedCols(cols);
      setKeepCols(cols);
      
      const initialRenameMap = {};
      cols.forEach(col => initialRenameMap[col] = col);
      setRenameMap(initialRenameMap);
    }
    e.target.value = null;
  };

  const toggleKeepCol = (col) => {
    setKeepCols(prev => 
      prev.includes(col) ? prev.filter(c => c !== col) : [...prev, col]
    );
  };

  const handleRenameChange = (oldName, newName) => {
    setRenameMap(prev => ({ ...prev, [oldName]: newName }));
  };

  const handleSort = () => {
    if (dragItem.current === null || dragOverItem.current === null) return;
    let _orderedCols = [...orderedCols];
    const draggedItemContent = _orderedCols.splice(dragItem.current, 1)[0];
    _orderedCols.splice(dragOverItem.current, 0, draggedItemContent);
    dragItem.current = null;
    dragOverItem.current = null;
    setOrderedCols(_orderedCols);
  };

  const handleExport = () => {
    if (data.length === 0) return;
    if (keepCols.length === 0) {
      alert("Выберите хотя бы одну колонку для экспорта!");
      return;
    }

    const colsToExport = orderedCols.filter(col => keepCols.includes(col));
    const finalHeaders = colsToExport.map(col => renameMap[col] || col);
    const csvRows = [finalHeaders.join(csvSeparator)];

    data.forEach(row => {
      const lineValues = colsToExport.map(col => {
        let rawVal = row[col];
        let val = "";

        if (rawVal instanceof Date) {
          const day = String(rawVal.getUTCDate()).padStart(2, '0');
          const month = String(rawVal.getUTCMonth() + 1).padStart(2, '0');
          const year = rawVal.getUTCFullYear();
          
          if (dateFormat === 'original' || dateFormat === 'DD.MM.YYYY') val = `${day}.${month}.${year}`;
          else if (dateFormat === 'DD-MM-YYYY') val = `${day}-${month}-${year}`;
          else if (dateFormat === 'DD/MM/YYYY') val = `${day}/${month}/${year}`;
          else if (dateFormat === 'YYYY.MM.DD') val = `${year}.${month}.${day}`;
          else if (dateFormat === 'YYYY-MM-DD') val = `${year}-${month}-${day}`;
          else if (dateFormat === 'YYYY/MM/DD') val = `${year}/${month}/${day}`;
          else if (dateFormat === 'MM.DD.YYYY') val = `${month}.${day}.${year}`;
          else if (dateFormat === 'MM-DD-YYYY') val = `${month}-${day}-${year}`;
          else if (dateFormat === 'MM/DD/YYYY') val = `${month}/${day}/${year}`;
        } 
        else {
          val = (rawVal !== undefined && rawVal !== null) ? rawVal.toString() : "";
          if (dateFormat !== 'original' && val) {
            const dmyMatch = val.match(/^(\d{1,2})[\.\-\/](\d{1,2})[\.\-\/](\d{2,4})$/);
            const ymdMatch = val.match(/^(\d{4})[\.\-\/](\d{1,2})[\.\-\/](\d{1,2})$/);
            
            let d, m, y, isDateString = false;
            
            if (dmyMatch) {
              d = dmyMatch[1].padStart(2, '0');
              m = dmyMatch[2].padStart(2, '0');
              y = dmyMatch[3];
              if (y.length === 2) y = parseInt(y) > 50 ? '19' + y : '20' + y;
              isDateString = true;
            } else if (ymdMatch) {
              y = ymdMatch[1];
              m = ymdMatch[2].padStart(2, '0');
              d = ymdMatch[3].padStart(2, '0');
              isDateString = true;
            }
            
            if (isDateString) {
              const checkDay = parseInt(d);
              const checkMonth = parseInt(m);
              if (checkDay >= 1 && checkDay <= 31 && checkMonth >= 1 && checkMonth <= 12) {
                if (dateFormat === 'DD.MM.YYYY') val = `${d}.${m}.${y}`;
                else if (dateFormat === 'DD-MM-YYYY') val = `${d}-${m}-${y}`;
                else if (dateFormat === 'DD/MM/YYYY') val = `${d}/${m}/${y}`;
                else if (dateFormat === 'YYYY.MM.DD') val = `${y}.${m}.${d}`;
                else if (dateFormat === 'YYYY-MM-DD') val = `${y}-${m}-${d}`;
                else if (dateFormat === 'YYYY/MM/DD') val = `${y}/${m}/${d}`;
                else if (dateFormat === 'MM.DD.YYYY') val = `${m}.${d}.${y}`;
                else if (dateFormat === 'MM-DD-YYYY') val = `${m}-${d}-${y}`;
                else if (dateFormat === 'MM/DD/YYYY') val = `${m}/${d}/${y}`;
              }
            }
          }
        }
        return val.replace(/\r?\n|\r/g, '').trim(); 
      });
      csvRows.push(lineValues.join(csvSeparator));
    });

    const csvString = '\uFEFF' + csvRows.join('\n');
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
    <div className="space-y-8">
      <div className="p-4 border-2 border-dashed border-gray-300 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 shadow-sm hover:bg-gray-50 dark:hover:bg-gray-700/50 transition-colors">
        <input 
          type="file" 
          accept=".xlsx, .xls" 
          onChange={handleFileUpload} 
          className="block w-full text-sm text-gray-500 dark:text-gray-400 file:mr-4 file:py-2 file:px-4 file:rounded file:border-0 file:text-sm file:font-semibold file:bg-blue-50 dark:file:bg-blue-950/40 file:text-blue-700 dark:file:text-blue-400 hover:file:bg-blue-100 dark:hover:file:bg-blue-900/40 cursor-pointer"
        />
      </div>

      {orderedCols.length > 0 && (
        <div className="space-y-6">
          <div className="space-y-4">
            <h2 className="text-xl font-semibold text-gray-700 dark:text-gray-200">Настройка экспорта</h2>
            <div className="border border-gray-200 dark:border-gray-800 p-4 rounded-lg bg-gray-50 dark:bg-gray-900/30 shadow-sm">
              <p className="text-sm text-gray-600 dark:text-gray-400 mb-4 font-medium">
                Снимите галочки с лишних колонок. Потяните строку мышкой за иконку слева, чтобы изменить порядок.
              </p>
              
              <div className="max-h-[500px] overflow-y-auto space-y-3 pr-2">
                {orderedCols.map((col, index) => {
                  const isKept = keepCols.includes(col);
                  return (
                    <div 
                      key={col} 
                      draggable
                      onDragStart={(e) => {
                        if (e.target.tagName.toLowerCase() === 'input') {
                          e.preventDefault();
                          return;
                        }
                        dragItem.current = index;
                      }}
                      onDragEnter={(e) => (dragOverItem.current = index)}
                      onDragEnd={handleSort}
                      onDragOver={(e) => e.preventDefault()}
                      className={`flex items-center space-x-4 p-3 border rounded-md shadow-sm transition-all cursor-move ${
                        isKept 
                          ? 'bg-white dark:bg-gray-800 border-gray-200 dark:border-gray-700 opacity-100 hover:border-blue-300 dark:hover:border-blue-700' 
                          : 'bg-gray-50 dark:bg-gray-900/40 border-gray-200 dark:border-gray-800 opacity-55'
                      }`}
                    >
                      <div className="text-gray-400 dark:text-gray-500 select-none" title="Потяните для сортировки">
                        <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 8h16M4 16h16"></path>
                        </svg>
                      </div>

                      <input 
                        type="checkbox" 
                        checked={isKept}
                        onChange={() => toggleKeepCol(col)}
                        className="rounded text-green-600 focus:ring-green-500 h-5 w-5 cursor-pointer dark:bg-gray-700 dark:border-gray-600"
                        title="Включить/выключить колонку"
                      />
                      <div className="flex-1 flex flex-col sm:flex-row sm:items-center sm:space-x-4">
                        <span className={`text-sm font-medium w-1/2 truncate cursor-text ${isKept ? 'text-gray-700 dark:text-gray-200' : 'text-gray-400 dark:text-gray-500 line-through'}`} title={col}>
                          {col}
                        </span>
                        <input 
                          type="text" 
                          value={renameMap[col] || ''}
                          onChange={(e) => handleRenameChange(col, e.target.value)}
                          className="w-full sm:w-1/2 p-2 border border-gray-300 dark:border-gray-700 rounded text-sm bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:border-green-500 focus:ring-1 focus:ring-green-500 disabled:bg-gray-100 dark:disabled:bg-gray-800 disabled:text-gray-400 dark:disabled:text-gray-500 cursor-text"
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

          <div className="pt-4 border-t border-gray-200 dark:border-gray-800 grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div className="flex items-center space-x-3">
              <label className="text-sm font-medium text-gray-700 dark:text-gray-300 whitespace-nowrap">Разделитель CSV:</label>
              <select 
                value={csvSeparator} 
                onChange={(e) => setCsvSeparator(e.target.value)}
                className="w-full p-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 text-sm focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none cursor-pointer"
              >
                <option value=";">Точка с запятой (;)</option>
                <option value=",">Запятая (,)</option>
              </select>
            </div>

            <div className="flex items-center space-x-3">
              <label className="text-sm font-medium text-gray-700 dark:text-gray-300 whitespace-nowrap">Формат даты:</label>
              <select 
                value={dateFormat} 
                onChange={(e) => setDateFormat(e.target.value)}
                className="w-full p-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 text-sm focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none cursor-pointer"
              >
                <option value="original">Без изменений</option>
                <option value="DD.MM.YYYY">ДД.ММ.ГГГГ (15.06.2026)</option>
                <option value="DD-MM-YYYY">ДД-ММ-ГГГГ (15-06-2026)</option>
                <option value="DD/MM/YYYY">ДД/ММ/ГГГГ (15/06/2026)</option>
                <option value="YYYY.MM.DD">ГГГГ.ММ.ДД (2026.06.15)</option>
                <option value="YYYY-MM-DD">ГГГГ-ММ-ДД (2026-06-15)</option>
                <option value="YYYY/MM/DD">ГГГГ/ММ/ДД (2026/06/15)</option>
                <option value="MM.DD.YYYY">ММ.ДД.ГГГГ (06.15.2026)</option>
                <option value="MM-DD-YYYY">ММ-ДД-ГГГГ (06-15-2026)</option>
                <option value="MM/DD/YYYY">ММ/ДД/ГГГГ (06/15/2026)</option>
              </select>
            </div>
            
            <button 
              onClick={handleExport}
              className="sm:col-span-2 w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-lg shadow hover:bg-blue-700 transition-colors focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 cursor-pointer mt-2"
            >
              Сохранить как CSV
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ==========================================
// КОМПОНЕНТ 2: CSV -> Excel
// ==========================================
function CsvToExcel() {
  const [csvFile, setCsvFile] = useState(null);
  const [delimiter, setDelimiter] = useState(';');

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (file) setCsvFile(file);
  };

  const handleExport = async () => {
    if (!csvFile) return;

    // Читаем текст из CSV файла
    const text = await csvFile.text();
    
    // Разбиваем на строки, очищаем пустые (если есть лишние переносы в конце)
    const rows = text.split(/\r?\n/).filter(row => row.trim() !== '');
    
    // Разбиваем строки на ячейки по выбранному разделителю
    const aoa = rows.map(row => row.split(delimiter));

    // Создаем лист и книгу Excel
    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    
    // Скачиваем готовый файл
    XLSX.writeFile(workbook, "converted_from_csv.xlsx");
  };

  return (
    <div className="space-y-8">
      <div className="p-4 border-2 border-dashed border-gray-300 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 shadow-sm hover:bg-gray-50 dark:hover:bg-gray-700/50 transition-colors">
        <input 
          type="file" 
          accept=".csv" 
          onChange={handleFileUpload} 
          className="block w-full text-sm text-gray-500 dark:text-gray-400 file:mr-4 file:py-2 file:px-4 file:rounded file:border-0 file:text-sm file:font-semibold file:bg-green-50 dark:file:bg-green-950/40 file:text-green-700 dark:file:text-green-400 hover:file:bg-green-100 dark:hover:file:bg-green-900/40 cursor-pointer"
        />
      </div>

      {csvFile && (
        <div className="pt-4 border-t border-gray-200 dark:border-gray-800 grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div className="flex items-center space-x-3">
            <label className="text-sm font-medium text-gray-700 dark:text-gray-300 whitespace-nowrap">Укажите разделитель CSV:</label>
            <select 
              value={delimiter} 
              onChange={(e) => setDelimiter(e.target.value)}
              className="w-full p-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 text-sm focus:border-green-500 focus:ring-1 focus:ring-green-500 outline-none cursor-pointer"
            >
              <option value=";">Точка с запятой (;)</option>
              <option value=",">Запятая (,)</option>
            </select>
          </div>

          <button 
            onClick={handleExport}
            className="sm:col-span-2 w-full bg-green-600 text-white font-bold py-3 px-4 rounded-lg shadow hover:bg-green-700 transition-colors focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 cursor-pointer mt-2"
          >
            Сконвертировать в Excel (.xlsx)
          </button>
        </div>
      )}
    </div>
  );
}

// ==========================================
// ОСНОВНОЙ КОМПОНЕНТ ПРИЛОЖЕНИЯ (С ТАБАМИ)
// ==========================================
export default function App() {
  const [activeTab, setActiveTab] = useState('excel-to-csv');

  return (
    <div className="max-w-3xl mx-auto p-6 font-sans text-gray-900 dark:text-gray-100 min-h-screen">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-8 space-y-4 sm:space-y-0">
        <h1 className="text-2xl font-bold text-gray-800 dark:text-gray-100">Конвертер данных</h1>
        
        {/* Переключатель вкладок */}
        <div className="flex space-x-2 bg-gray-100 dark:bg-gray-800 p-1 rounded-lg">
          <button 
            onClick={() => setActiveTab('excel-to-csv')}
            className={`px-4 py-2 text-sm font-semibold rounded-md transition-all ${
              activeTab === 'excel-to-csv' 
                ? 'bg-white dark:bg-gray-700 text-blue-600 dark:text-blue-400 shadow-sm' 
                : 'text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-200 hover:bg-gray-200 dark:hover:bg-gray-700/50'
            }`}
          >
            Excel ➡️ CSV
          </button>
          <button 
            onClick={() => setActiveTab('csv-to-excel')}
            className={`px-4 py-2 text-sm font-semibold rounded-md transition-all ${
              activeTab === 'csv-to-excel' 
                ? 'bg-white dark:bg-gray-700 text-green-600 dark:text-green-400 shadow-sm' 
                : 'text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-200 hover:bg-gray-200 dark:hover:bg-gray-700/50'
            }`}
          >
            CSV ➡️ Excel
          </button>
        </div>
      </div>

      {/* Отрисовка активного режима */}
      {activeTab === 'excel-to-csv' ? <ExcelToCsv /> : <CsvToExcel />}
    </div>
  );
}