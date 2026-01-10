import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { createCustomer, createUserTracuu } from '../services/api';
import './ImportExcel.css';

const ImportExcel = () => {
  const [file, setFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [results, setResults] = useState(null);
  const [errors, setErrors] = useState([]);
  const [progress, setProgress] = useState({ current: 0, total: 0 });

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      setResults(null);
      setErrors([]);
    }
  };

  const readExcelFile = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(firstSheet);
          resolve(jsonData);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  const processImport = async () => {
    if (!file) {
      alert('Vui lòng chọn file Excel!');
      return;
    }

    setProcessing(true);
    setResults(null);
    setErrors([]);
    setProgress({ current: 0, total: 0 });

    try {
      const jsonData = await readExcelFile(file);
      
      // Tìm cột mã đối tượng (có thể là "mã đối tượng", "ma_doi_tuong", "ma_dt", hoặc cột đầu tiên)
      const firstRow = jsonData[0];
      let maDoiTuongKey = null;
      
      if (firstRow) {
        // Tìm key có chứa "mã" hoặc "ma"
        maDoiTuongKey = Object.keys(firstRow).find(key => 
          key.toLowerCase().includes('mã') || 
          key.toLowerCase().includes('ma') ||
          key.toLowerCase().includes('code') ||
          key.toLowerCase().includes('id')
        );
        
        // Nếu không tìm thấy, lấy cột đầu tiên
        if (!maDoiTuongKey) {
          maDoiTuongKey = Object.keys(firstRow)[0];
        }
      }

      if (!maDoiTuongKey) {
        throw new Error('Không tìm thấy cột mã đối tượng trong file Excel!');
      }

      const maDoiTuongList = jsonData
        .map(row => {
          const value = row[maDoiTuongKey];
          return value ? String(value).trim() : null;
        })
        .filter(val => val && val !== '');

      if (maDoiTuongList.length === 0) {
        throw new Error('Không tìm thấy mã đối tượng nào trong file Excel!');
      }

      // Khởi tạo progress với tổng số
      setProgress({ current: 0, total: maDoiTuongList.length });

      const successList = [];
      const errorList = [];

      // Xử lý từng mã đối tượng
      for (let i = 0; i < maDoiTuongList.length; i++) {
        const ma_dt = maDoiTuongList[i];
        
        // Cập nhật progress
        setProgress({ current: i + 1, total: maDoiTuongList.length });
        
        try {
          // Bước 1: Tạo danh mục khách hàng
          const customerResult = await createCustomer(ma_dt);
          
          if (!customerResult.success) {
            const errorMsg = customerResult.error || 
              (customerResult.data ? (typeof customerResult.data === 'string' ? customerResult.data : JSON.stringify(customerResult.data)) : 'Lỗi không xác định');
            errorList.push({
              ma_dt,
              step: 'Tạo danh mục khách hàng',
              error: errorMsg
            });
            continue;
          }

          // Kiểm tra response từ API tạo danh mục khách hàng có lỗi không
          if (customerResult.data && typeof customerResult.data === 'object') {
            if (customerResult.data.error || customerResult.data.message) {
              const errorMsg = customerResult.data.error || customerResult.data.message;
              errorList.push({
                ma_dt,
                step: 'Tạo danh mục khách hàng',
                error: errorMsg
              });
              continue;
            }
          }

          // Bước 2: Tạo user tra cứu
          const userResult = await createUserTracuu(ma_dt);
          
          if (!userResult.success) {
            const errorMsg = userResult.error || 
              (userResult.data ? (typeof userResult.data === 'string' ? userResult.data : JSON.stringify(userResult.data)) : 'Lỗi không xác định');
            errorList.push({
              ma_dt,
              step: 'Tạo user tra cứu',
              error: errorMsg
            });
            continue;
          }

          // Kiểm tra response từ API tạo user
          if (userResult.data) {
            // Trường hợp thành công: có trường "ok"
            if (userResult.data.ok) {
              successList.push({
                ma_dt,
                message: userResult.data.ok
              });
            } 
            // Trường hợp lỗi: có trường "error" hoặc "message" với nội dung lỗi
            else if (userResult.data.error || (userResult.data.message && !userResult.data.message.includes('thành công'))) {
              const errorMsg = userResult.data.error || userResult.data.message;
              errorList.push({
                ma_dt,
                step: 'Tạo user tra cứu',
                error: errorMsg
              });
            } 
            // Trường hợp response là string
            else if (typeof userResult.data === 'string') {
              if (userResult.data.toLowerCase().includes('lỗi') || userResult.data.toLowerCase().includes('error')) {
                errorList.push({
                  ma_dt,
                  step: 'Tạo user tra cứu',
                  error: userResult.data
                });
              } else {
                successList.push({
                  ma_dt,
                  message: userResult.data
                });
              }
            }
            // Trường hợp khác, coi như thành công nếu không có dấu hiệu lỗi
            else {
              successList.push({
                ma_dt,
                message: 'Thành công'
              });
            }
          } else {
            // Không có data, coi như thành công
            successList.push({
              ma_dt,
              message: 'Thành công'
            });
          }

        } catch (error) {
          errorList.push({
            ma_dt,
            step: 'Xử lý',
            error: error.message
          });
        }
      }

      setResults({
        total: maDoiTuongList.length,
        success: successList.length,
        failed: errorList.length
      });
      setErrors(errorList);

    } catch (error) {
      alert('Lỗi: ' + error.message);
    } finally {
      setProcessing(false);
    }
  };

  const exportErrorsToExcel = () => {
    if (errors.length === 0) {
      alert('Không có lỗi để xuất!');
      return;
    }

    const ws = XLSX.utils.json_to_sheet(errors);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Danh sách lỗi');
    
    const fileName = `Danh_sach_loi_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  return (
    <div className="import-excel-container">
      <div className="import-excel-card">
        <h1 className="import-excel-title">Import Excel - Tạo Khách Hàng & User Tra Cứu</h1>
        
        <div className="import-excel-upload">
          <label className="file-input-label">
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileChange}
              disabled={processing}
              className="file-input"
            />
            <span className="file-input-text">
              {file ? file.name : 'Chọn file Excel (.xlsx, .xls)'}
            </span>
            <span className="file-input-button">Chọn File</span>
          </label>
        </div>

        <button
          onClick={processImport}
          disabled={!file || processing}
          className="import-button"
        >
          {processing ? 'Đang xử lý...' : 'Bắt đầu Import'}
        </button>

        {processing && (
          <div className="processing-indicator">
            <div className="spinner"></div>
            <p>Đang xử lý dữ liệu, vui lòng đợi...</p>
            {progress.total > 0 && (
              <div className="progress-container">
                <div className="progress-info">
                  <span>Đang xử lý: {progress.current}/{progress.total} đối tượng</span>
                  <span className="progress-percentage">
                    {Math.round((progress.current / progress.total) * 100)}%
                  </span>
                </div>
                <div className="progress-bar-wrapper">
                  <div 
                    className="progress-bar" 
                    style={{ width: `${(progress.current / progress.total) * 100}%` }}
                  ></div>
                </div>
              </div>
            )}
          </div>
        )}

        {results && (
          <div className="results-container">
            <div className="results-summary">
              <div className="result-item success">
                <span className="result-label">Tổng số:</span>
                <span className="result-value">{results.total}</span>
              </div>
              <div className="result-item success">
                <span className="result-label">Thành công:</span>
                <span className="result-value">{results.success}</span>
              </div>
              <div className="result-item failed">
                <span className="result-label">Thất bại:</span>
                <span className="result-value">{results.failed}</span>
              </div>
            </div>

            {errors.length > 0 && (
              <div className="errors-section">
                <div className="errors-header">
                  <h3>Danh sách lỗi ({errors.length})</h3>
                  <button onClick={exportErrorsToExcel} className="export-button">
                    Xuất Excel
                  </button>
                </div>
                <div className="errors-table-container">
                  <table className="errors-table">
                    <thead>
                      <tr>
                        <th>STT</th>
                        <th>Mã đối tượng</th>
                        <th>Bước lỗi</th>
                        <th>Lý do</th>
                      </tr>
                    </thead>
                    <tbody>
                      {errors.map((error, index) => (
                        <tr key={index}>
                          <td>{index + 1}</td>
                          <td>{error.ma_dt}</td>
                          <td>{error.step}</td>
                          <td className="error-reason">{error.error}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {errors.length === 0 && results.success > 0 && (
              <div className="success-message">
                <p>✓ Tất cả các mã đối tượng đã được xử lý thành công!</p>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default ImportExcel;
