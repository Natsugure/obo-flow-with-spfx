import express, { Request, Response } from 'express';
import { getAuthUrl, getTokenFromCode } from './auth';
import { createSubscription, getEmailMessage, extractMessageIdFromResource, listSubscriptions } from './graphClient';
import { saveUserToken, addSubscription, getAllUsers, getUserToken, saveEmailMessage, getAllEmailMessages } from './db';
import { startRenewalService } from './renewalService';
import { NotificationPayload, StoredEmailMessage } from './types';
import { configureAxios } from './config/axios';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
app.use(express.json());

configureAxios();

/**
 * ホームページ
 */
app.get('/', (req: Request, res: Response) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Microsoft Graph Subscription Demo</title>
      <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 50px auto; padding: 20px; }
        h1 { color: #0078d4; }
        a { display: inline-block; padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px; }
        a:hover { background: #106ebe; }
      </style>
    </head>
    <body>
      <h1>📧 Microsoft Graph Subscription Demo</h1>
      <p>このデモでは、Microsoft Graphのサブスクリプション機能を使用して、受信トレイの変更通知を受け取ります。</p>
      <a href="/auth/signin">🔐 サインインしてサブスクライブ</a>
      <br><br>
      <a href="/emails">📧 受信メール一覧</a>
      <br><br>
      <a href="/status">📊 ステータス確認</a>
    </body>
    </html>
  `);
});

/**
 * サインイン開始
 */
app.get('/auth/signin', async (req: Request, res: Response) => {
  try {
    const authUrl = await getAuthUrl();
    res.redirect(authUrl);
  } catch (error) {
    console.error('認証URL生成エラー:', error);
    res.status(500).send('認証URLの生成に失敗しました');
  }
});

/**
 * 認証コールバック
 */
app.get('/auth/callback', async (req: Request, res: Response) => {
  const code = req.query.code as string;
  
  if (!code) {
    return res.status(400).send('認証コードがありません');
  }
  
  try {
    // トークンを取得
    const tokenData = await getTokenFromCode(code);
    const userId = tokenData.account.homeAccountId;
    
    // トークンを保存
    saveUserToken(userId, tokenData);
    
    // サブスクリプションを作成
    const subscription = await createSubscription(tokenData.accessToken, userId);
    addSubscription(userId, subscription);
    
    res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>認証成功</title>
        <style>
          body { font-family: Arial, sans-serif; max-width: 800px; margin: 50px auto; padding: 20px; }
          h1 { color: #107c10; }
          .info { background: #f3f2f1; padding: 15px; border-radius: 5px; margin: 10px 0; }
          .info strong { color: #0078d4; }
          a { display: inline-block; margin-top: 20px; padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px; }
        </style>
      </head>
      <body>
        <h1>✅ 認証成功！</h1>
        <div class="info">
          <p><strong>ユーザーID:</strong> ${userId}</p>
          <p><strong>サブスクリプションID:</strong> ${subscription.id}</p>
          <p><strong>有効期限:</strong> ${new Date(subscription.expirationDateTime).toLocaleString('ja-JP')}</p>
        </div>
        <p>✉️ これで受信トレイの変更通知を受信できます。</p>
        <p>📬 自分宛てにメールを送信してテストしてください。</p>
        <a href="/emails">📧 受信メール一覧</a>
        <a href="/status" style="margin-left: 10px; background: #6c757d;">📊 ステータス確認</a>
        <a href="/" style="margin-left: 10px; background: #6c757d;">🏠 ホームに戻る</a>
      </body>
      </html>
    `);
  } catch (error) {
    console.error('認証エラー:', error);
    const errorMessage = error instanceof Error ? error.message : '不明なエラー';
    res.status(500).send(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>認証エラー</title>
      </head>
      <body>
        <h1>❌ 認証に失敗しました</h1>
        <p>エラー: ${errorMessage}</p>
        <a href="/">ホームに戻る</a>
      </body>
      </html>
    `);
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
 * ステータス確認
 */
app.get('/status', (req: Request, res: Response) => {
  const users = getAllUsers();
  
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>ステータス確認</title>
      <style>
        body { font-family: Arial, sans-serif; max-width: 1000px; margin: 50px auto; padding: 20px; }
        h1 { color: #0078d4; }
        .user-card { background: #f3f2f1; padding: 20px; margin: 20px 0; border-radius: 5px; }
        .user-card h2 { color: #323130; margin-top: 0; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #0078d4; color: white; }
        .expires-soon { color: #d13438; font-weight: bold; }
        .expires-ok { color: #107c10; }
        a { display: inline-block; margin-top: 20px; padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px; }
        .check-btn { background: #107c10; border: none; color: white; padding: 10px 15px; border-radius: 5px; cursor: pointer; font-size: 14px; }
        .check-btn:hover { background: #0e6b0e; }
        .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.5); }
        .modal-content { background-color: #fefefe; margin: 5% auto; padding: 20px; border-radius: 8px; width: 90%; max-width: 800px; max-height: 80vh; overflow-y: auto; }
        .close { color: #aaa; float: right; font-size: 28px; font-weight: bold; cursor: pointer; }
        .close:hover { color: #000; }
        .json-display { background: #f5f5f5; padding: 15px; border-radius: 5px; overflow-x: auto; white-space: pre-wrap; font-family: monospace; font-size: 12px; }
        .loading { color: #0078d4; font-style: italic; }
        .error { color: #d13438; padding: 10px; background: #fef0f0; border-radius: 5px; }
      </style>
      <script>
        async function checkSubscriptions(userId) {
          const modal = document.getElementById('subscriptionModal');
          const content = document.getElementById('subscriptionContent');
          modal.style.display = 'block';
          content.innerHTML = '<p class="loading">🔄 サブスクリプション情報を取得中...</p>';

          try {
            const response = await fetch('/api/subscriptions/' + encodeURIComponent(userId));
            const data = await response.json();

            if (!response.ok) {
              throw new Error(data.error || 'サブスクリプション取得エラー');
            }

            content.innerHTML = '<h2>✅ サブスクリプション情報</h2>' +
              '<p><strong>ユーザーID:</strong> ' + data.userId + '</p>' +
              '<p><strong>サブスクリプション数:</strong> ' + data.count + '</p>' +
              '<p><strong>取得日時:</strong> ' + new Date(data.timestamp).toLocaleString('ja-JP') + '</p>' +
              '<h3>Response Body:</h3>' +
              '<div class="json-display">' + JSON.stringify(data.subscriptions, null, 2) + '</div>';
          } catch (error) {
            content.innerHTML = '<div class="error"><strong>❌ エラー:</strong> ' + error.message + '</div>';
          }
        }

        function closeModal() {
          document.getElementById('subscriptionModal').style.display = 'none';
        }

        window.onclick = function(event) {
          const modal = document.getElementById('subscriptionModal');
          if (event.target === modal) {
            modal.style.display = 'none';
          }
        }
      </script>
    </head>
    <body>
      <h1>📊 現在のステータス</h1>
      <p>登録ユーザー数: ${users.length}</p>
  `;
  
  if (users.length === 0) {
    html += '<p>まだユーザーが登録されていません。</p>';
  } else {
    users.forEach(user => {
      const tokenExpiry = new Date(user.expiresOn);
      const tokenTimeLeft = tokenExpiry.getTime() - Date.now();
      const tokenMinutesLeft = Math.floor(tokenTimeLeft / 60000);
      const tokenClass = tokenMinutesLeft < 10 ? 'expires-soon' : 'expires-ok';
      
      html += `
        <div class="user-card">
          <h2>👤 ユーザー: ${user.userId}</h2>
          <p><strong>トークン有効期限:</strong>
            <span class="${tokenClass}">
              ${tokenExpiry.toLocaleString('ja-JP')}
              (残り約${tokenMinutesLeft}分)
            </span>
          </p>
          <div style="margin: 15px 0;">
            <button class="check-btn" onclick="checkSubscriptions('${user.userId}')">
              🔍 サブスクリプションの確認
            </button>
          </div>
          <h3>📋 ローカルDBのサブスクリプション:</h3>
      `;
      
      if (user.subscriptions.length === 0) {
        html += '<p>サブスクリプションがありません</p>';
      } else {
        html += `
          <table>
            <tr>
              <th>ID</th>
              <th>リソース</th>
              <th>有効期限</th>
              <th>残り時間</th>
            </tr>
        `;
        
        user.subscriptions.forEach(sub => {
          const subExpiry = new Date(sub.expirationDateTime);
          const subTimeLeft = subExpiry.getTime() - Date.now();
          const subHoursLeft = Math.floor(subTimeLeft / 3600000);
          const subClass = subHoursLeft < 1 ? 'expires-soon' : 'expires-ok';
          
          html += `
            <tr>
              <td>${sub.id.substring(0, 8)}...</td>
              <td>${sub.resource}</td>
              <td>${subExpiry.toLocaleString('ja-JP')}</td>
              <td class="${subClass}">約${subHoursLeft}時間</td>
            </tr>
          `;
        });
        
        html += '</table>';
      }
      
      html += '</div>';
    });
  }
  
  html += `
      <!-- モーダル -->
      <div id="subscriptionModal" class="modal">
        <div class="modal-content">
          <span class="close" onclick="closeModal()">&times;</span>
          <div id="subscriptionContent"></div>
        </div>
      </div>

      <a href="/">🏠 ホームに戻る</a>
      <a href="/emails" style="margin-left: 10px; padding: 10px 20px; background: #107c10; color: white; text-decoration: none; border-radius: 5px; display: inline-block;">📧 受信メール一覧</a>
      <script>
        // 30秒ごとに自動更新
        setTimeout(() => location.reload(), 30000);
      </script>
    </body>
    </html>
  `;

  res.send(html);
});

/**
 * 受信メール一覧表示
 */
app.get('/emails', (req: Request, res: Response) => {
  const emails = getAllEmailMessages();

  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>受信メール一覧</title>
      <style>
        body { font-family: Arial, sans-serif; max-width: 1200px; margin: 50px auto; padding: 20px; background: #f5f5f5; }
        h1 { color: #0078d4; }
        .email-card { background: white; padding: 20px; margin: 20px 0; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .email-header { display: flex; justify-content: space-between; align-items: start; margin-bottom: 15px; }
        .email-subject { font-size: 20px; font-weight: bold; color: #323130; margin: 0 0 10px 0; }
        .email-from { color: #605e5c; margin: 5px 0; }
        .email-from strong { color: #323130; }
        .email-time { color: #8a8886; font-size: 14px; }
        .email-preview { color: #605e5c; margin: 15px 0; padding: 15px; background: #faf9f8; border-radius: 4px; }
        .email-body { margin-top: 15px; padding: 15px; background: #faf9f8; border-radius: 4px; max-height: 300px; overflow-y: auto; }
        .email-meta { display: flex; gap: 20px; margin-top: 15px; padding-top: 15px; border-top: 1px solid #edebe9; }
        .email-meta-item { font-size: 14px; color: #605e5c; }
        .email-meta-item strong { color: #323130; }
        .badge { display: inline-block; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; }
        .badge-attachment { background: #d13438; color: white; }
        .no-emails { text-align: center; padding: 40px; color: #8a8886; }
        a.button { display: inline-block; margin-top: 20px; padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 5px; }
        a.button:hover { background: #106ebe; }
        .stats { background: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; display: flex; gap: 30px; }
        .stat-item { flex: 1; }
        .stat-number { font-size: 32px; font-weight: bold; color: #0078d4; }
        .stat-label { color: #605e5c; font-size: 14px; }
        .toggle-body { cursor: pointer; color: #0078d4; font-size: 14px; margin-top: 10px; }
        .toggle-body:hover { text-decoration: underline; }
      </style>
      <script>
        function toggleBody(id) {
          const element = document.getElementById('body-' + id);
          const toggle = document.getElementById('toggle-' + id);
          if (element.style.display === 'none') {
            element.style.display = 'block';
            toggle.textContent = '▼ 本文を非表示';
          } else {
            element.style.display = 'none';
            toggle.textContent = '▶ 本文を表示';
          }
        }
      </script>
    </head>
    <body>
      <h1>📧 受信メール一覧</h1>

      <div class="stats">
        <div class="stat-item">
          <div class="stat-number">${emails.length}</div>
          <div class="stat-label">受信メール数</div>
        </div>
        <div class="stat-item">
          <div class="stat-number">${emails.filter(e => e.hasAttachments).length}</div>
          <div class="stat-label">添付ファイル付き</div>
        </div>
      </div>
  `;

  if (emails.length === 0) {
    html += `
      <div class="no-emails">
        <h2>📭 まだメールが届いていません</h2>
        <p>サブスクリプションが作成されたら、自分宛てにメールを送信してテストしてください。</p>
      </div>
    `;
  } else {
    emails.forEach((email, index) => {
      const receivedDate = new Date(email.receivedDateTime);
      const notificationDate = new Date(email.notificationReceived);

      html += `
        <div class="email-card">
          <div class="email-header">
            <div style="flex: 1;">
              <h2 class="email-subject">${escapeHtml(email.subject || '(件名なし)')}</h2>
              <div class="email-from">
                <strong>差出人:</strong> ${escapeHtml(email.from)}
                &lt;${escapeHtml(email.fromAddress)}&gt;
              </div>
            </div>
            <div class="email-time">
              ${receivedDate.toLocaleString('ja-JP')}
            </div>
          </div>

          <div class="email-preview">
            <strong>プレビュー:</strong> ${escapeHtml(email.bodyPreview)}
          </div>

          <div class="toggle-body" id="toggle-${index}" onclick="toggleBody(${index})">
            ▶ 本文を表示
          </div>

          <div class="email-body" id="body-${index}" style="display: none;">
            ${email.bodyContent}
          </div>

          <div class="email-meta">
            <div class="email-meta-item">
              <strong>メッセージID:</strong> ${email.id.substring(0, 20)}...
            </div>
            <div class="email-meta-item">
              <strong>通知受信:</strong> ${notificationDate.toLocaleString('ja-JP')}
            </div>
            ${email.hasAttachments ? '<span class="badge badge-attachment">📎 添付ファイルあり</span>' : ''}
          </div>
        </div>
      `;
    });
  }

  html += `
      <a href="/" class="button">🏠 ホームに戻る</a>
      <a href="/status" class="button" style="margin-left: 10px; background: #6c757d;">📊 ステータス確認</a>
      <script>
        // 30秒ごとに自動更新
        setTimeout(() => location.reload(), 30000);
      </script>
    </body>
    </html>
  `;

  res.send(html);
});

/**
 * HTMLエスケープ関数
 */
function escapeHtml(text: string): string {
  const map: { [key: string]: string } = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, (m) => map[m]);
}

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