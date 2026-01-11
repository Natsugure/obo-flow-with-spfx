import axios from 'axios';

/**
 * axiosのグローバル設定
 */
export function configureAxios(): void {
  // デフォルトタイムアウト
  axios.defaults.timeout = 30000;
  
  // リクエストインターセプター
  axios.interceptors.request.use(
    (config) => {
      console.log(`→ ${config.method?.toUpperCase()} ${config.url}`);
      return config;
    },
    (error) => {
      console.error('リクエストエラー:', error);
      return Promise.reject(error);
    }
  );
  
  // レスポンスインターセプター
  axios.interceptors.response.use(
    (response) => {
      console.log(`← ${response.status} ${response.config.url}`);
      return response;
    },
    (error) => {
      if (axios.isAxiosError(error)) {
        if (error.response) {
          console.error(`← ${error.response.status} ${error.config?.url}`);
        } else if (error.request) {
          console.error('← レスポンスなし:', error.config?.url);
        }
      }
      return Promise.reject(error);
    }
  );
  
  console.log('✓ Axios設定完了');
}