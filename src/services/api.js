// API configuration
const API_BASE_URL = "https://1702325579.minvoice.com.vn/api";
const AUTH_TOKEN = "O87316arj5+Od3Fqyy5hzdBfIuPk73eKqpAzBSvv8sY=";
const MA_DVCS = "VP";
const LANGUAGE = "vi";

// Common headers
const getHeaders = () => ({
  Accept: "*/*",
  "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7,fr-FR;q=0.6,fr;q=0.5",
  Authorization: `Bear ${AUTH_TOKEN}`,
  "Cache-Control": "no-cache",
  Connection: "keep-alive",
  "Content-type": "application/json",
  Origin: "https://1702325579.minvoice.com.vn",
  Pragma: "no-cache",
  Referer: "https://1702325579.minvoice.com.vn/",
  "Sec-Fetch-Dest": "empty",
  "Sec-Fetch-Mode": "cors",
  "Sec-Fetch-Site": "same-origin",
  "User-Agent":
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36",
  "sec-ch-ua":
    '"Google Chrome";v="143", "Chromium";v="143", "Not A(Brand";v="24"',
  "sec-ch-ua-mobile": "?0",
  "sec-ch-ua-platform": '"macOS"',
});

// API 1: Tạo danh mục khách hàng
export const createCustomer = async (ma_dt) => {
  try {
    const response = await fetch(`${API_BASE_URL}/System/Save`, {
      method: "POST",
      headers: getHeaders(),
      body: JSON.stringify({
        windowid: "WIN00009",
        editmode: 1,
        data: [
          {
            ma_dvcs: MA_DVCS,
            ma_dt: ma_dt,
            ms_thue: "",
            dt_me_id: "",
            ten_dt: ma_dt,
            email: "",
            dai_dien: "",
            dia_chi: ma_dt,
            dien_thoai: "",
            dien_giai: "",
            fax: "",
            details: [
              {
                tab_id: "TAB00014",
                tab_table: "dmngh",
                data: [
                  {
                    id: Date.now(),
                    ma_dvcs: MA_DVCS,
                    idx: 1,
                    so_tk: "",
                    dmngh_id: null,
                    dmdt_id: null,
                  },
                ],
              },
            ],
          },
        ],
      }),
    });

    const data = await response.json();
    return { success: response.ok, data };
  } catch (error) {
    return { success: false, error: error.message };
  }
};

// API 2: Tạo user tra cứu
export const createUserTracuu = async (ma_dt) => {
  try {
    // mst phải là mã số thuế cố định (1702325579) theo curl command
    const mst = "1702325579";

    const response = await fetch(`${API_BASE_URL}/Invoice/CreateUser_tracuu`, {
      method: "POST",
      headers: getHeaders(),
      body: JSON.stringify({
        mst: mst,
        ma_dt: ma_dt,
        username: ma_dt,
        password: ma_dt,
        email: "",
      }),
    });

    const data = await response.json();
    return { success: response.ok, data };
  } catch (error) {
    return { success: false, error: error.message };
  }
};
