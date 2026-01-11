declare global {
  namespace NodeJS {
    interface ProcessEnv {
      CLIENT_ID: string;
      CLIENT_SECRET: string;
      TENANT_ID: string;
      REDIRECT_URI: string;
      NOTIFICATION_URL: string;
      PORT?: string;
      CLIENT_STATE_SECRET: string;
      ENCRYPTION_KEY?: string;
      WINDOWS_OPENSSL_PATH?: string;
    }
  }
}

export {};