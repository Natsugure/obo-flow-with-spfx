import { AccountInfo } from '@azure/msal-node';
import { TokenData, UserTokenData, SubscriptionData, Subscription, StoredEmailMessage } from './types';

// 簡易的なインメモリストレージ（本番環境ではPostgreSQL/MongoDBなどを使用）
const users = new Map<string, UserTokenData>();
const emailMessages = new Map<string, StoredEmailMessage>();

/**
 * ユーザートークンクラス
 */
class UserToken implements UserTokenData {
  userId: string;
  accessToken: string;
  refreshToken: string;
  expiresOn: Date;
  account: AccountInfo;
  subscriptions: SubscriptionData[];

  constructor(
    userId: string,
    accessToken: string,
    refreshToken: string,
    expiresOn: Date,
    account: AccountInfo
  ) {
    this.userId = userId;
    this.accessToken = accessToken;
    this.refreshToken = refreshToken;
    this.expiresOn = expiresOn;
    this.account = account;
    this.subscriptions = [];
  }
}

/**
 * ユーザートークンを保存
 */
export function saveUserToken(userId: string, tokenData: TokenData): void {
  const userToken = new UserToken(
    userId,
    tokenData.accessToken,
    tokenData.refreshToken,
    tokenData.expiresOn,
    tokenData.account
  );
  users.set(userId, userToken);
  console.log(`ユーザー ${userId} のトークンを保存しました`);
}

/**
 * ユーザートークンを取得
 */
export function getUserToken(userId: string): UserTokenData | undefined {
  return users.get(userId);
}

/**
 * ユーザートークンを更新
 */
export function updateUserToken(userId: string, tokenData: TokenData): void {
  const user = users.get(userId);
  if (user) {
    user.accessToken = tokenData.accessToken;
    user.refreshToken = tokenData.refreshToken || user.refreshToken;
    user.expiresOn = tokenData.expiresOn;
    console.log(`ユーザー ${userId} のトークンを更新しました`);
  }
}

/**
 * サブスクリプションを追加
 */
export function addSubscription(userId: string, subscription: Subscription): void {
  const user = users.get(userId);
  if (user) {
    user.subscriptions.push({
      id: subscription.id,
      expirationDateTime: subscription.expirationDateTime,
      resource: subscription.resource
    });
    console.log(`ユーザー ${userId} にサブスクリプション ${subscription.id} を追加しました`);
  }
}

/**
 * サブスクリプションを更新
 */
export function updateSubscriptionExpiration(
  userId: string, 
  subscriptionId: string, 
  expirationDateTime: string
): void {
  const user = users.get(userId);
  if (user) {
    const subscription = user.subscriptions.find(sub => sub.id === subscriptionId);
    if (subscription) {
      subscription.expirationDateTime = expirationDateTime;
      console.log(`サブスクリプション ${subscriptionId} の有効期限を更新しました`);
    }
  }
}

/**
 * 全ユーザーを取得
 */
export function getAllUsers(): UserTokenData[] {
  return Array.from(users.values());
}

/**
 * ユーザーを削除
 */
export function deleteUser(userId: string): boolean {
  return users.delete(userId);
}

/**
 * メールメッセージを保存
 */
export function saveEmailMessage(message: StoredEmailMessage): void {
  emailMessages.set(message.id, message);
  console.log(`メールメッセージ ${message.id} を保存しました (件名: ${message.subject})`);
}

/**
 * メールメッセージを取得
 */
export function getEmailMessage(messageId: string): StoredEmailMessage | undefined {
  return emailMessages.get(messageId);
}

/**
 * 全メールメッセージを取得（新しい順）
 */
export function getAllEmailMessages(): StoredEmailMessage[] {
  const messages = Array.from(emailMessages.values());
  return messages.sort((a, b) =>
    new Date(b.notificationReceived).getTime() - new Date(a.notificationReceived).getTime()
  );
}

/**
 * ユーザーIDでメールメッセージを取得（新しい順）
 */
export function getEmailMessagesByUserId(userId: string): StoredEmailMessage[] {
  const messages = Array.from(emailMessages.values()).filter(msg => msg.userId === userId);
  return messages.sort((a, b) =>
    new Date(b.notificationReceived).getTime() - new Date(a.notificationReceived).getTime()
  );
}

/**
 * メールメッセージを削除
 */
export function deleteEmailMessage(messageId: string): boolean {
  return emailMessages.delete(messageId);
}

/**
 * 全メールメッセージを削除
 */
export function clearAllEmailMessages(): void {
  emailMessages.clear();
  console.log('全てのメールメッセージを削除しました');
}