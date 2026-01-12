import express, { Request, Response } from 'express';
import { acquireTokenOnBehalfOf } from './auth';
import { createSubscription, getEmailMessage, extractMessageIdFromResource, listSubscriptions } from './graphClient';
import { saveUserToken, addSubscription, getAllUsers, getUserToken, saveEmailMessage, getAllEmailMessages } from './db';
import { startRenewalService } from './renewalService';
import { NotificationPayload, StoredEmailMessage } from './types';
import { configureAxios } from './config/axios';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
app.use(express.json());

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*'); // 本番環境では特定のドメインに制限
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
  
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  
  next();
});

configureAxios();

app.post('/api/subscribe', async (req: Request, res: Response) => {
  const authHeader = req.headers.authorization;
  if(!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: '認証トークンがありません' });
  }

  const userToken = authHeader.substring(7);

  try {
    const oboTokenResponse = await acquireTokenOnBehalfOf(userToken);

    if (!oboTokenResponse) {
      return res.status(401).json({ 
        error: 'トークンの取得に失敗しました',
        details: 'OBO認証が失敗しました'
      });
    }

    const subscription = await createSubscription(
      oboTokenResponse.accessToken,
      oboTokenResponse.account.homeAccountId,
    );

    // TODO: Azure SQL Databaseへの保存ロジックを実装
    // saveToDatabase()

    res.json({
      userId: oboTokenResponse.account.homeAccountId,
      subscriptionId: subscription.id,
      expirationDateTime: subscription.expirationDateTime,
      resource: subscription.resource,
    })

  } catch (error) {
    console.error('サブスクリプション作成エラー:', error);
    res.status(500).json({ 
      error: 'サブスクリプションの作成に失敗しました',
      details: error instanceof Error ? error.message : '不明なエラー'
    });
  }
});

/**
 * 通知を受信するエンドポイント
 */
app.post('/api/notifications', async (req: Request, res: Response) => {
  // 検証リクエスト
  const validationToken = req.query.validationToken as string;
  if (validationToken) {
    console.log('✓ 検証トークンを受信:', validationToken);
    return res.type('text/plain').send(validationToken);
  }

  // 実際の通知
  const payload: NotificationPayload = req.body;
  const notifications = payload.value;

  console.log('\n📬 通知を受信:', new Date().toISOString());
  console.log(JSON.stringify(notifications, null, 2));

  // 非同期処理を並列実行
  const processingPromises = notifications.map(async (notification) => {
    console.log(`  📝 変更タイプ: ${notification.changeType}`);
    console.log(`  📂 リソース: ${notification.resource}`);
    console.log(`  🔐 クライアント状態: ${notification.clientState}`);

    // clientStateの検証
    if (notification.clientState !== process.env.CLIENT_STATE_SECRET) {
      console.warn('  ⚠️  無効なclientState');
      return;
    }

    try {
      // メッセージIDを抽出
      const messageId = extractMessageIdFromResource(notification.resource);
      console.log(`  📧 メッセージID: ${messageId}`);

      // サブスクリプションIDからユーザーを特定
      const users = getAllUsers();
      const user = users.find(u =>
        u.subscriptions.some(sub => sub.id === notification.subscriptionId)
      );

      if (!user) {
        console.warn(`  ⚠️  サブスクリプション ${notification.subscriptionId} に対応するユーザーが見つかりません`);
        return;
      }

      // Graph APIからメール詳細を取得
      console.log(`  🔍 メール詳細を取得中...`);
      const emailData = await getEmailMessage(user.accessToken, messageId);

      // データベースに保存
      const storedMessage: StoredEmailMessage = {
        id: emailData.id,
        userId: user.userId,
        subscriptionId: notification.subscriptionId,
        subject: emailData.subject,
        from: emailData.from.emailAddress.name || emailData.from.emailAddress.address,
        fromAddress: emailData.from.emailAddress.address,
        bodyPreview: emailData.bodyPreview,
        bodyContent: emailData.body.content,
        receivedDateTime: emailData.receivedDateTime,
        notificationReceived: new Date().toISOString(),
        hasAttachments: emailData.hasAttachments,
        webLink: emailData.webLink
      };

      saveEmailMessage(storedMessage);
      console.log(`  ✅ メール保存完了: ${emailData.subject}`);
    } catch (error) {
      console.error(`  ❌ メール取得エラー:`, error);
      if (error instanceof Error) {
        console.error(`     詳細: ${error.message}`);
      }
    }
  });

  // 全ての通知処理を待機（ただしレスポンスはすぐ返す）
  Promise.all(processingPromises).catch(err => {
    console.error('通知処理でエラーが発生しました:', err);
  });

  // 必ず202を返す（Microsoft Graphの要件）
  res.status(202).send();
});



/**
 * ユーザーのサブスクリプション情報を取得
 */
app.get('/api/subscriptions/:userId', async (req: Request, res: Response) => {
  const userId = req.params.userId;

  // userIdが配列の場合はエラー
  if (Array.isArray(userId)) {
    return res.status(400).json({ error: '不正なユーザーIDです' });
  }

  try {
    const user = getUserToken(userId);

    if (!user) {
      return res.status(404).json({ error: 'ユーザーが見つかりません' });
    }

    // Graph APIからサブスクリプション一覧を取得
    const subscriptions = await listSubscriptions(user.accessToken);

    res.json({
      userId: user.userId,
      subscriptions: subscriptions,
      count: subscriptions.length,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    console.error('サブスクリプション取得エラー:', error);
    const errorMessage = error instanceof Error ? error.message : '不明なエラー';
    res.status(500).json({
      error: 'サブスクリプション情報の取得に失敗しました',
      details: errorMessage
    });
  }
});

/**
 * ヘルスチェック
 */
app.get('/health', (req: Request, res: Response) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    users: getAllUsers().length,
    emails: getAllEmailMessages().length
  });
});

/**
 * サーバー起動
 */
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('=====================================');
  console.log('🚀 サーバーが起動しました');
  console.log(`📍 URL: http://localhost:${PORT}`);
  console.log(`🌐 Notification URL: ${process.env.NOTIFICATION_URL}`);
  console.log('=====================================\n');
  
  // 更新サービスを開始
  startRenewalService();
});

// グレースフルシャットダウン
process.on('SIGINT', () => {
  console.log('\n🛑 サーバーをシャットダウンしています...');
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('\n🛑 サーバーをシャットダウンしています...');
  process.exit(0);
});