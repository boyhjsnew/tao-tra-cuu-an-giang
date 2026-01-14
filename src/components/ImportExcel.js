import React, { useState } from "react";
import * as XLSX from "xlsx";
import { createCustomer, createUserTracuu } from "../services/api";
import "./ImportExcel.css";

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
          const workbook = XLSX.read(data, {
            type: "array",
            cellDates: false,
            cellNF: false,
            cellText: false,
          });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

          // Kiểm tra range của sheet
          if (!firstSheet["!ref"]) {
            throw new Error("Sheet không có dữ liệu!");
          }

          const range = XLSX.utils.decode_range(firstSheet["!ref"]);
          const totalRowsInSheet = range.e.r + 1;

          console.log("Sheet range:", firstSheet["!ref"]);
          console.log("Total rows in sheet (from range):", totalRowsInSheet);

          // Đọc tất cả dữ liệu - cách chuẩn nhất
          // Không dùng blankrows: true vì nó có thể gây vấn đề với file lớn
          const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
            defval: null, // Giá trị mặc định cho ô trống
            blankrows: false, // Bỏ qua dòng trống
            raw: false, // Chuyển đổi giá trị
          });

          console.log("Total rows read (json):", jsonData.length);

          // Nếu số dòng đọc được ít hơn đáng kể so với range, có thể có vấn đề
          // Thử đọc lại bằng cách khác - đọc dạng array rồi convert
          if (
            jsonData.length < totalRowsInSheet * 0.9 &&
            totalRowsInSheet > 100
          ) {
            console.warn(
              "Phát hiện số dòng đọc được ít hơn range, thử phương pháp đọc khác..."
            );

            // Đọc dạng array (bao gồm cả header)
            const arrayData = XLSX.utils.sheet_to_json(firstSheet, {
              header: 1,
              defval: null,
              blankrows: false,
            });

            console.log("Total rows read (array):", arrayData.length);

            if (arrayData.length > jsonData.length + 1) {
              // Convert từ array sang object
              if (arrayData.length > 0 && arrayData[0]) {
                // Xử lý header - đảm bảo không có undefined
                const headers = arrayData[0].map((h, index) => {
                  if (h === null || h === undefined || h === "") {
                    return `Column_${index + 1}`;
                  }
                  return String(h).trim() || `Column_${index + 1}`;
                });

                console.log("Headers:", headers);

                const result = [];

                for (let i = 1; i < arrayData.length; i++) {
                  if (!arrayData[i]) continue;

                  const row = {};
                  headers.forEach((header, index) => {
                    const value = arrayData[i][index];
                    row[header] =
                      value !== undefined && value !== null ? value : null;
                  });
                  // Chỉ thêm row nếu có ít nhất 1 giá trị không null
                  if (Object.values(row).some((v) => v !== null && v !== "")) {
                    result.push(row);
                  }
                }

                console.log(
                  "Total rows after array conversion:",
                  result.length
                );
                resolve(result);
                return;
              }
            }
          }

          resolve(jsonData);
        } catch (error) {
          console.error("Error reading Excel:", error);
          reject(error);
        }
      };
      reader.onerror = (error) => {
        console.error("FileReader error:", error);
        reject(error);
      };
      reader.readAsArrayBuffer(file);
    });
  };

  const processImport = async () => {
    if (!file) {
      alert("Vui lòng chọn file Excel!");
      return;
    }

    setProcessing(true);
    setResults(null);
    setErrors([]);
    setProgress({ current: 0, total: 0 });

    try {
      const jsonData = await readExcelFile(file);

      console.log("Total rows from Excel:", jsonData.length);

      if (!jsonData || jsonData.length === 0) {
        throw new Error("File Excel không có dữ liệu!");
      }

      // Tìm cột mã đối tượng (có thể là "mã đối tượng", "ma_doi_tuong", "ma_dt", hoặc cột đầu tiên)
      const firstRow = jsonData[0];
      let maDoiTuongKey = null;

      if (!firstRow || typeof firstRow !== "object") {
        throw new Error("Dòng đầu tiên của file Excel không hợp lệ!");
      }

      // Lấy danh sách các key có giá trị
      const availableKeys = Object.keys(firstRow).filter(
        (key) => key !== null && key !== undefined
      );

      if (availableKeys.length === 0) {
        throw new Error("Không tìm thấy cột nào trong file Excel!");
      }

      console.log("Available columns:", availableKeys);

      // Tìm key có chứa "mã" hoặc "ma"
      maDoiTuongKey = availableKeys.find((key) => {
        const lowerKey = String(key).toLowerCase();
        return (
          lowerKey.includes("mã") ||
          lowerKey.includes("ma") ||
          lowerKey.includes("code") ||
          lowerKey.includes("id") ||
          lowerKey.includes("mã đối tượng") ||
          lowerKey.includes("ma doi tuong")
        );
      });

      // Nếu không tìm thấy, lấy cột đầu tiên
      if (!maDoiTuongKey) {
        maDoiTuongKey = availableKeys[0];
        console.warn(
          'Không tìm thấy cột "mã đối tượng", sử dụng cột đầu tiên:',
          maDoiTuongKey
        );
      }

      if (!maDoiTuongKey) {
        throw new Error("Không tìm thấy cột mã đối tượng trong file Excel!");
      }

      console.log("Column key found:", maDoiTuongKey);
      console.log("Sample first row:", firstRow);

      // Lọc và map tất cả các dòng, bao gồm cả dòng có giá trị null/undefined
      const maDoiTuongList = jsonData
        .map((row, index) => {
          // Kiểm tra row có tồn tại và có key không
          if (!row || typeof row !== "object") {
            console.warn(`Row ${index + 1} không hợp lệ:`, row);
            return null;
          }

          // Kiểm tra key có tồn tại trong row không
          if (!(maDoiTuongKey in row)) {
            console.warn(`Row ${index + 1} không có key "${maDoiTuongKey}"`);
            return null;
          }

          const value = row[maDoiTuongKey];

          // Chuyển đổi tất cả giá trị thành string và trim
          if (value !== null && value !== undefined && value !== "") {
            try {
              const strValue = String(value).trim();
              return strValue !== "" ? strValue : null;
            } catch (error) {
              console.warn(
                `Lỗi khi chuyển đổi giá trị ở row ${index + 1}:`,
                error
              );
              return null;
            }
          }
          return null;
        })
        .filter((val) => val !== null && val !== "" && val !== undefined);

      console.log("Total valid rows after filter:", maDoiTuongList.length);
      console.log("Sample data (first 5):", maDoiTuongList.slice(0, 5));

      if (maDoiTuongList.length === 0) {
        throw new Error("Không tìm thấy mã đối tượng nào trong file Excel!");
      }

      // Khởi tạo progress với tổng số
      setProgress({ current: 0, total: maDoiTuongList.length });

      const successList = [];
      const errorList = [];

      // Xử lý từng mã đối tượng
      for (let i = 0; i < maDoiTuongList.length; i++) {
        const ma_dt = maDoiTuongList[i];

        // Kiểm tra ma_dt có hợp lệ không
        if (!ma_dt || ma_dt === "undefined" || ma_dt === "null") {
          errorList.push({
            ma_dt: ma_dt || `Row ${i + 1}`,
            step: "Xử lý",
            error: "Mã đối tượng không hợp lệ",
          });
          setProgress({ current: i + 1, total: maDoiTuongList.length });
          continue;
        }

        // Cập nhật progress
        setProgress({ current: i + 1, total: maDoiTuongList.length });

        try {
          // Bước 1: Tạo danh mục khách hàng (độc lập)
          let customerSuccess = false;
          const customerResult = await createCustomer(ma_dt);

          if (!customerResult || !customerResult.success) {
            let errorMsg = "Lỗi không xác định";

            if (customerResult) {
              if (customerResult.error) {
                errorMsg = String(customerResult.error);
              } else if (customerResult.data) {
                if (typeof customerResult.data === "string") {
                  errorMsg = customerResult.data;
                } else if (typeof customerResult.data === "object") {
                  errorMsg = JSON.stringify(customerResult.data);
                }
              }
            }

            errorList.push({
              ma_dt: String(ma_dt),
              step: "Tạo danh mục khách hàng",
              error: errorMsg,
            });
          } else {
            // Kiểm tra response từ API tạo danh mục khách hàng có lỗi không
            if (
              customerResult.data &&
              typeof customerResult.data === "object"
            ) {
              if (customerResult.data.error || customerResult.data.message) {
                const errorMsg =
                  customerResult.data.error || customerResult.data.message;
                errorList.push({
                  ma_dt,
                  step: "Tạo danh mục khách hàng",
                  error: errorMsg,
                });
              } else {
                customerSuccess = true;
              }
            } else {
              customerSuccess = true;
            }
          }

          // Bước 2: Tạo user tra cứu (độc lập, luôn thực hiện dù tạo khách hàng thành công hay thất bại)
          const userResult = await createUserTracuu(ma_dt);

          // Xử lý kết quả tạo user tra cứu (độc lập)
          if (!userResult || !userResult.success) {
            let errorMsg = "Lỗi không xác định";

            if (userResult) {
              if (userResult.error) {
                errorMsg = String(userResult.error);
              } else if (userResult.data) {
                if (typeof userResult.data === "string") {
                  errorMsg = userResult.data;
                } else if (typeof userResult.data === "object") {
                  errorMsg = JSON.stringify(userResult.data);
                }
              }
            }

            errorList.push({
              ma_dt: String(ma_dt),
              step: "Tạo user tra cứu",
              error: errorMsg,
            });
          } else {
            // Kiểm tra response từ API tạo user
            if (userResult.data) {
              // Trường hợp thành công: có trường "ok"
              if (userResult.data.ok) {
                successList.push({
                  ma_dt,
                  message: `Tạo user tra cứu thành công. ${
                    customerSuccess
                      ? "Tạo khách hàng thành công."
                      : "Tạo khách hàng thất bại."
                  }`,
                });
              }
              // Trường hợp lỗi: có trường "error" hoặc "message" với nội dung lỗi
              else if (
                userResult.data.error ||
                (userResult.data.message &&
                  !userResult.data.message.includes("thành công"))
              ) {
                const errorMsg =
                  userResult.data.error || userResult.data.message;
                errorList.push({
                  ma_dt,
                  step: "Tạo user tra cứu",
                  error: errorMsg,
                });
              }
              // Trường hợp response là string
              else if (typeof userResult.data === "string") {
                if (
                  userResult.data.toLowerCase().includes("lỗi") ||
                  userResult.data.toLowerCase().includes("error")
                ) {
                  errorList.push({
                    ma_dt,
                    step: "Tạo user tra cứu",
                    error: userResult.data,
                  });
                } else {
                  successList.push({
                    ma_dt,
                    message: `Tạo user tra cứu thành công. ${
                      customerSuccess
                        ? "Tạo khách hàng thành công."
                        : "Tạo khách hàng thất bại."
                    }`,
                  });
                }
              }
              // Trường hợp khác, coi như thành công nếu không có dấu hiệu lỗi
              else {
                successList.push({
                  ma_dt,
                  message: `Tạo user tra cứu thành công. ${
                    customerSuccess
                      ? "Tạo khách hàng thành công."
                      : "Tạo khách hàng thất bại."
                  }`,
                });
              }
            } else {
              // Không có data, coi như thành công
              successList.push({
                ma_dt,
                message: `Tạo user tra cứu thành công. ${
                  customerSuccess
                    ? "Tạo khách hàng thành công."
                    : "Tạo khách hàng thất bại."
                }`,
              });
            }
          }
        } catch (error) {
          const errorMessage =
            error && error.message
              ? String(error.message)
              : error
              ? String(error)
              : "Lỗi không xác định";
          errorList.push({
            ma_dt: ma_dt ? String(ma_dt) : `Row ${i + 1}`,
            step: "Xử lý",
            error: errorMessage,
          });
        }
      }

      setResults({
        total: maDoiTuongList.length,
        success: successList.length,
        failed: errorList.length,
      });
      setErrors(errorList);
    } catch (error) {
      alert("Lỗi: " + error.message);
    } finally {
      setProcessing(false);
    }
  };

  const exportErrorsToExcel = () => {
    if (errors.length === 0) {
      alert("Không có lỗi để xuất!");
      return;
    }

    const ws = XLSX.utils.json_to_sheet(errors);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Danh sách lỗi");

    const fileName = `Danh_sach_loi_${
      new Date().toISOString().split("T")[0]
    }.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  return (
    <div className="import-excel-container">
      <div className="import-excel-card">
        <h1 className="import-excel-title">
          Import Excel - Tạo Khách Hàng & User Tra Cứu
        </h1>

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
              {file ? file.name : "Chọn file Excel (.xlsx, .xls)"}
            </span>
            <span className="file-input-button">Chọn File</span>
          </label>
        </div>

        <button
          onClick={processImport}
          disabled={!file || processing}
          className="import-button"
        >
          {processing ? "Đang xử lý..." : "Bắt đầu Import"}
        </button>

        {processing && (
          <div className="processing-indicator">
            <div className="spinner"></div>
            <p>Đang xử lý dữ liệu, vui lòng đợi...</p>
            {progress.total > 0 && (
              <div className="progress-container">
                <div className="progress-info">
                  <span>
                    Đang xử lý: {progress.current}/{progress.total} đối tượng
                  </span>
                  <span className="progress-percentage">
                    {Math.round((progress.current / progress.total) * 100)}%
                  </span>
                </div>
                <div className="progress-bar-wrapper">
                  <div
                    className="progress-bar"
                    style={{
                      width: `${(progress.current / progress.total) * 100}%`,
                    }}
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
                  <button
                    onClick={exportErrorsToExcel}
                    className="export-button"
                  >
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
                          <td>{error.ma_dt || "N/A"}</td>
                          <td>{error.step || "N/A"}</td>
                          <td className="error-reason">
                            {error.error || "Lỗi không xác định"}
                          </td>
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
