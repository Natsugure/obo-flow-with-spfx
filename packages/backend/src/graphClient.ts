import { Client } from '@microsoft/microsoft-graph-client';
import axios, { AxiosInstance, AxiosError } from 'axios';
import { Subscription, EmailMessage } from './types';

/**
 * Graph APIクライアントを取得
 */
function getGraphClient(accessToken: string): Client {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
}

/**
 * Axiosインスタンスを作成（直接Graph APIを呼ぶ場合用）
 */
export function createGraphAxiosClient(accessToken: string): AxiosInstance {
  return axios.create({
    baseURL: 'https://graph.microsoft.com/v1.0',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    timeout: 30000 // 30秒タイムアウト
  });
}

/**
 * サブスクリプションの作成
 */
export async function createSubscription(
  accessToken: string, 
  userId: string
): Promise<Subscription> {
  const client = getGraphClient(accessToken);
  
  const subscription = {
    changeType: 'created',
    notificationUrl: process.env.NOTIFICATION_URL,
    resource: '/me/mailFolders/inbox/messages',
    expirationDateTime: getExpirationDateTime(4230), // 最大4230分（約3日）
    clientState: process.env.CLIENT_STATE_SECRET
  };
  
  try {
    const result = await client.api('/subscriptions').post(subscription);
    console.log('サブスクリプション作成成功:', result);
    return result as Subscription;
  } catch (error) {
    handleGraphError('サブスクリプション作成', error);
    throw error;
  }
}

/**
 * サブスクリプションの更新
 */
export async function updateSubscription(
  accessToken: string, 
  subscriptionId: string
): Promise<Subscription> {
  const client = getGraphClient(accessToken);
  
  const update = {
    expirationDateTime: getExpirationDateTime(4230)
  };
  
  try {
    const result = await client
      .api(`/subscriptions/${subscriptionId}`)
      .patch(update);
    console.log('サブスクリプション更新成功:', result);
    return result as Subscription;
  } catch (error) {
    handleGraphError('サブスクリプション更新', error);
    throw error;
  }
}

/**
 * サブスクリプションの削除
 */
export async function deleteSubscription(
  accessToken: string, 
  subscriptionId: string
): Promise<void> {
  const client = getGraphClient(accessToken);
  
  try {
    await client.api(`/subscriptions/${subscriptionId}`).delete();
    console.log('サブスクリプション削除成功');
  } catch (error) {
    handleGraphError('サブスクリプション削除', error);
    throw error;
  }
}

/**
 * axiosを使ってサブスクリプションを直接作成（代替実装例）
 */
export async function createSubscriptionWithAxios(
  accessToken: string,
  userId: string
): Promise<Subscription> {
  const axiosClient = createGraphAxiosClient(accessToken);
  
  const subscription = {
    changeType: 'created',
    notificationUrl: process.env.NOTIFICATION_URL,
    resource: '/me/mailFolders/inbox/messages',
    expirationDateTime: getExpirationDateTime(4230),
    clientState: process.env.CLIENT_STATE_SECRET
  };
  
  try {
    const response = await axiosClient.post<Subscription>('/subscriptions', subscription);
    console.log('サブスクリプション作成成功 (axios):', response.data);
    return response.data;
  } catch (error) {
    handleAxiosError('サブスクリプション作成 (axios)', error);
    throw error;
  }
}

/**
 * サブスクリプション一覧を取得（axios実装例）
 */
export async function listSubscriptions(accessToken: string): Promise<Subscription[]> {
  const axiosClient = createGraphAxiosClient(accessToken);
  
  try {
    const response = await axiosClient.get<{ value: Subscription[] }>('/subscriptions');
    return response.data.value;
  } catch (error) {
    handleAxiosError('サブスクリプション一覧取得', error);
    throw error;
  }
}

/**
 * 有効期限の計算（現在時刻から指定分後）
 */
function getExpirationDateTime(minutes: number): string {
  const expiration = new Date();
  expiration.setMinutes(expiration.getMinutes() + minutes);
  return expiration.toISOString();
}

/**
 * Graph APIエラーハンドリング
 */
function handleGraphError(operation: string, error: unknown): void {
  console.error(`${operation}エラー:`, error);
  
  if (axios.isAxiosError(error)) {
    const axiosError = error as AxiosError;
    if (axiosError.response) {
      console.error('ステータスコード:', axiosError.response.status);
      console.error('レスポンスデータ:', axiosError.response.data);
    }
  }
}

/**
 * Axiosエラーハンドリング
 */
function handleAxiosError(operation: string, error: unknown): void {
  if (axios.isAxiosError(error)) {
    const axiosError = error as AxiosError;

    console.error(`${operation}エラー:`);

    if (axiosError.response) {
      // サーバーがエラーレスポンスを返した
      console.error('ステータスコード:', axiosError.response.status);
      console.error('レスポンスデータ:', JSON.stringify(axiosError.response.data, null, 2));

      // 特定のエラーコードに応じた処理
      switch (axiosError.response.status) {
        case 401:
          console.error('認証エラー: トークンが無効または期限切れです');
          throw new Error('REAUTH_REQUIRED');
        case 403:
          console.error('権限エラー: 必要な権限がありません');
          break;
        case 404:
          console.error('リソースが見つかりません');
          break;
        case 429:
          console.error('レート制限超過: しばらく待ってから再試行してください');
          break;
        default:
          console.error('予期しないエラーが発生しました');
      }
    } else if (axiosError.request) {
      // リクエストは送信されたがレスポンスがない
      console.error('レスポンスなし:', axiosError.message);
    } else {
      // リクエスト設定中にエラーが発生
      console.error('リクエストエラー:', axiosError.message);
    }
  } else {
    console.error(`${operation}で予期しないエラー:`, error);
  }
}

/**
 * メールメッセージの詳細を取得
 */
export async function getEmailMessage(
  accessToken: string,
  messageId: string
): Promise<EmailMessage> {
  const client = getGraphClient(accessToken);

  try {
    const message = await client
      .api(`/me/messages/${messageId}`)
      .select('id,subject,bodyPreview,body,from,receivedDateTime,hasAttachments,webLink')
      .get();

    return message as EmailMessage;
  } catch (error) {
    handleGraphError('メールメッセージ取得', error);
    throw error;
  }
}

/**
 * リソースURLからメッセージIDを抽出
 * 例: "Users/xxx/Messages/yyy" -> "yyy"
 */
export function extractMessageIdFromResource(resource: string): string {
  const parts = resource.split('/');
  return parts[parts.length - 1];
}