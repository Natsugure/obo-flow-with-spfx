import axios from 'axios';
import { refreshAccessToken } from './auth';
import { updateSubscription } from './graphClient';
import { getAllUsers, updateUserToken, updateSubscriptionExpiration } from './db';
import { UserTokenData } from './types';

/**
 * トークンの更新が必要かチェック
 */
function needsTokenRenewal(expiresOn: Date): boolean {
  const now = new Date();
  const expiration = new Date(expiresOn);
  const timeUntilExpiry = expiration.getTime() - now.getTime();
  
  // 有効期限の10分前になったら更新
  return timeUntilExpiry < 10 * 60 * 1000;
}

/**
 * サブスクリプションの更新が必要かチェック
 */
function needsSubscriptionRenewal(expirationDateTime: string): boolean {
  const now = new Date();
  const expiration = new Date(expirationDateTime);
  const timeUntilExpiry = expiration.getTime() - now.getTime();
  
  // 有効期限の1時間前になったら更新
  return timeUntilExpiry < 60 * 60 * 1000;
}

/**
 * 定期的な更新処理
 */
async function renewalTask(): Promise<void> {
  console.log('=== 更新タスク実行中 ===', new Date().toISOString());
  const users = getAllUsers();
  
  if (users.length === 0) {
    console.log('登録ユーザーがいません');
    return;
  }
  
  for (const user of users) {
    try {
      // トークンの更新チェック
      if (needsTokenRenewal(user.expiresOn)) {
        console.log(`ユーザー ${user.userId} のトークンを更新中...`);
        
        try {
          const newTokens = await refreshAccessToken(user.account);
          updateUserToken(user.userId, newTokens);
          console.log(`✓ ユーザー ${user.userId} のトークン更新成功`);
          
          // 更新したトークンを使用
          user.accessToken = newTokens.accessToken;
        } catch (tokenError) {
          if (tokenError instanceof Error && tokenError.message === 'REAUTH_REQUIRED') {
            console.error(`❌ ユーザー ${user.userId} は再認証が必要です`);
            await notifyUserForReauth(user.userId);
            continue; // このユーザーのサブスクリプション更新はスキップ
          }
          throw tokenError;
        }
      }
      
      // サブスクリプションの更新チェック
      for (const subscription of user.subscriptions) {
        if (needsSubscriptionRenewal(subscription.expirationDateTime)) {
          console.log(`サブスクリプション ${subscription.id} を更新中...`);
          
          try {
            const updated = await updateSubscription(user.accessToken, subscription.id);
            updateSubscriptionExpiration(user.userId, subscription.id, updated.expirationDateTime);
            console.log(`✓ サブスクリプション ${subscription.id} 更新成功`);
          } catch (subError) {
            if (axios.isAxiosError(subError) && subError.response?.status === 404) {
              console.warn(`⚠️  サブスクリプション ${subscription.id} が見つかりません（削除済み？）`);
              // TODO: DBから削除する処理
            } else if (axios.isAxiosError(subError) && subError.response?.status === 401) {
              console.error(`❌ サブスクリプション更新で認証エラー - ユーザー ${user.userId} は再認証が必要です`);
              await notifyUserForReauth(user.userId);
              break; // このユーザーの残りのサブスクリプション更新もスキップ
            } else {
              throw subError;
            }
          }
        }
      }
    } catch (error) {
      console.error(`❌ ユーザー ${user.userId} の更新で予期しないエラー:`, error);
      
      if (axios.isAxiosError(error)) {
        const axiosError = error;
        if (axiosError.response) {
          console.error('ステータスコード:', axiosError.response.status);
          console.error('エラー詳細:', axiosError.response.data);
        }
      }
    }
  }
  
  console.log('=== 更新タスク完了 ===\n');
}

/**
 * 再認証通知（実装例 - Teams通知を想定）
 */
async function notifyUserForReauth(userId: string): Promise<void> {
  console.log(`📧 【通知】ユーザー ${userId} に再認証を依頼`);
  
  // TODO: 実際の通知実装
  // 例1: Teams Adaptive Cardで通知
  // await sendTeamsAdaptiveCard(userId, {
  //   title: '再認証が必要です',
  //   body: 'メール通知サービスとの接続が切れました。再度ログインしてください。',
  //   actions: [{ type: 'openUrl', title: '再ログイン', url: 'https://your-app.com/auth/signin' }]
  // });
  
  // 例2: メール通知
  // await sendEmail({
  //   to: await getUserEmail(userId),
  //   subject: '再認証のお願い',
  //   body: '接続が切れたため、再度認証が必要です。'
  // });
  
  // 例3: Webhookで外部システムに通知
  // try {
  //   await axios.post('https://your-webhook-url.com/notify', {
  //     userId,
  //     event: 'reauth_required',
  //     timestamp: new Date().toISOString()
  //   });
  // } catch (error) {
  //   console.error('Webhook通知エラー:', error);
  // }
}

/**
 * 更新サービスを開始
 */
let renewalIntervalId: NodeJS.Timeout | null = null;

export function startRenewalService(): void {
  console.log('🔄 更新サービスを開始しました（5分間隔）');
  
  // 初回実行
  renewalTask().catch(error => {
    console.error('初回更新タスクエラー:', error);
  });
  
  // 定期実行（5分ごと）
  renewalIntervalId = setInterval(() => {
    renewalTask().catch(error => {
      console.error('更新タスクエラー:', error);
    });
  }, 5 * 60 * 1000);
}

/**
 * 更新サービスを停止
 */
export function stopRenewalService(): void {
  if (renewalIntervalId) {
    clearInterval(renewalIntervalId);
    renewalIntervalId = null;
    console.log('🛑 更新サービスを停止しました');
  }
}

/**
 * 即座に更新タスクを実行（テスト用）
 */
export async function runRenewalTaskNow(): Promise<void> {
  await renewalTask();
}