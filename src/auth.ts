import * as msal from '@azure/msal-node';
import { AccountInfo, AuthenticationResult } from '@azure/msal-node';
import dotenv from 'dotenv';
import { TokenData } from './types';

dotenv.config();

// MSAL設定
const msalConfig: msal.Configuration = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel: msal.LogLevel, message: string, containsPii: boolean) {
        if (!containsPii) {
          console.log(message);
        }
      },
      piiLoggingEnabled: false,
      logLevel: msal.LogLevel.Verbose,
    }
  }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

/**
 * 認証URLを生成
 */
export async function getAuthUrl(): Promise<string> {
  const authCodeUrlParameters: msal.AuthorizationUrlRequest = {
    scopes: ['User.Read', 'Mail.Read', 'offline_access'],
    redirectUri: process.env.REDIRECT_URI,
  };
  
  return await cca.getAuthCodeUrl(authCodeUrlParameters);
}

/**
 * 認証コードからトークンを取得
 */
export async function getTokenFromCode(code: string): Promise<TokenData> {
  const tokenRequest: msal.AuthorizationCodeRequest = {
    code: code,
    scopes: ['User.Read', 'Mail.Read', 'offline_access'],
    redirectUri: process.env.REDIRECT_URI,
  };
  
  const response: AuthenticationResult = await cca.acquireTokenByCode(tokenRequest);
  
  if (!response.accessToken || !response.account) {
    throw new Error('トークンまたはアカウント情報の取得に失敗しました');
  }
  
  // MSALはリフレッシュトークンを直接返さない
  // トークンキャッシュに保存される
  return {
    accessToken: response.accessToken,
    refreshToken: '', // MSALがキャッシュで管理するため空文字
    expiresOn: response.expiresOn || new Date(),
    account: response.account
  };
}

/**
 * サイレントトークン取得（リフレッシュトークンの代わり）
 */
export async function refreshAccessToken(
  account: AccountInfo
): Promise<TokenData> {
  const silentRequest: msal.SilentFlowRequest = {
    account: account,
    scopes: ['User.Read', 'Mail.Read'],
    forceRefresh: false // キャッシュを優先、期限切れなら自動更新
  };
  
  try {
    const response: AuthenticationResult = await cca.acquireTokenSilent(silentRequest);
    
    if (!response.accessToken || !response.account) {
      throw new Error('トークンの更新に失敗しました');
    }
    
    return {
      accessToken: response.accessToken,
      refreshToken: '', // MSALがキャッシュで管理
      expiresOn: response.expiresOn || new Date(),
      account: response.account
    };
  } catch (error) {
    console.error('サイレントトークン取得エラー:', error);
    
    // キャッシュにトークンがない、または期限切れで更新できない
    if (error instanceof msal.InteractionRequiredAuthError) {
      throw new Error('REAUTH_REQUIRED');
    }
    
    throw new Error('REAUTH_REQUIRED');
  }
}

/**
 * トークンを強制的に更新
 */
export async function forceRefreshAccessToken(
  account: AccountInfo
): Promise<TokenData> {
  const silentRequest: msal.SilentFlowRequest = {
    account: account,
    scopes: ['User.Read', 'Mail.Read'],
    forceRefresh: true // 強制的に新しいトークンを取得
  };
  
  try {
    const response: AuthenticationResult = await cca.acquireTokenSilent(silentRequest);
    
    if (!response.accessToken || !response.account) {
      throw new Error('トークンの更新に失敗しました');
    }
    
    return {
      accessToken: response.accessToken,
      refreshToken: '',
      expiresOn: response.expiresOn || new Date(),
      account: response.account
    };
  } catch (error) {
    console.error('強制トークン更新エラー:', error);
    throw new Error('REAUTH_REQUIRED');
  }
}