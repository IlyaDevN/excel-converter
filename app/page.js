'use client';

import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';

export default function ExcelConverter() {
  const [data, setData] = useState([]);
  
  // orderedCols хранит актуальный порядок всех колонок
  const [orderedCols, setOrderedCols] = useState([]);
  
  const [keepCols, setKeepCols] = useState([]);
  const [renameMap, setRenameMap] = useState({});
  
  // Состояние для разделителя
  const [csvSeparator, setCsvSeparator] = useState(';');

  // Рефы для Drag & Drop
  const dragItem = useRef(null);
  const dragOverItem = useRef(null);

  // Чтение и парсинг Excel файла
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // ДОБАВЛЕН ПАРАМЕТР raw: false
    // Он заставляет парсер брать видимый текст ячейки (формат даты), а не внутреннее число Excel
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
  };

  // Переключение чекбокса (включить/выключить колонку)
  const toggleKeepCol = (col) => {
    setKeepCols(prev => 
      prev.includes(col) ? prev.filter(c => c !== col) : [...prev, col]
    );
  };

  const handleRenameChange = (oldName, newName) => {
    setRenameMap(prev => ({ ...prev, [oldName]: newName }));
  };

  // Логика сортировки при перетаскивании
  const handleSort = () => {
    if (dragItem.current === null || dragOverItem.current === null) return;
    
    let _orderedCols = [...orderedCols];
    const draggedItemContent = _orderedCols.splice(dragItem.current, 1)[0];
    _orderedCols.splice(dragOverItem.current, 0, draggedItemContent);
    
    dragItem.current = null;
    dragOverItem.current = null;
    setOrderedCols(_orderedCols);
  };

  // Генерация и скачивание CSV
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
        let val = (row[col] || "").toString();
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

      {orderedCols.length > 0 && (
        <div className="space-y-6">
          <div className="space-y-4">
            <h2 className="text-xl font-semibold text-gray-700">Настройка экспорта</h2>
            <div className="border p-4 rounded-lg bg-gray-50 shadow-sm">
              <p className="text-sm text-gray-600 mb-4 font-medium">
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
                      className={`flex items-center space-x-4 p-3 bg-white border rounded-md shadow-sm transition-all cursor-move hover:border-blue-300 ${isKept ? 'opacity-100' : 'opacity-60 bg-gray-50'}`}
                    >
                      <div className="text-gray-400 select-none" title="Потяните для сортировки">
                        <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 8h16M4 16h16"></path>
                        </svg>
                      </div>

                      <input 
                        type="checkbox" 
                        checked={isKept}
                        onChange={() => toggleKeepCol(col)}
                        className="rounded text-green-600 focus:ring-green-500 h-5 w-5 cursor-pointer"
                        title="Включить/выключить колонку"
                      />
                      <div className="flex-1 flex flex-col sm:flex-row sm:items-center sm:space-x-4">
                        <span className={`text-sm font-medium w-1/2 truncate cursor-text ${isKept ? 'text-gray-700' : 'text-gray-400 line-through'}`} title={col}>
                          {col}
                        </span>
                        <input 
                          type="text" 
                          value={renameMap[col] || ''}
                          onChange={(e) => handleRenameChange(col, e.target.value)}
                          className="w-full sm:w-1/2 p-2 border border-gray-300 rounded text-sm focus:border-green-500 focus:ring-1 focus:ring-green-500 disabled:bg-gray-100 disabled:text-gray-400 cursor-text"
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

          <div className="pt-4 border-t border-gray-200 space-y-4">
            <div className="flex items-center space-x-3">
              <label className="text-sm font-medium text-gray-700">Разделитель CSV:</label>
              <select 
                value={csvSeparator} 
                onChange={(e) => setCsvSeparator(e.target.value)}
                className="p-2 border border-gray-300 rounded-md shadow-sm bg-white text-sm focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none cursor-pointer"
              >
                <option value=";">Точка с запятой (;)</option>
                <option value=",">Запятая (,)</option>
              </select>
            </div>
            
            <button 
              onClick={handleExport}
              className="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-lg shadow hover:bg-blue-700 transition-colors focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 cursor-pointer"
            >
              Сохранить как CSV
            </button>
          </div>
        </div>
      )}
    </div>
  );
}